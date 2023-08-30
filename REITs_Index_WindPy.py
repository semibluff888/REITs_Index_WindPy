# 本代码是下面jupyter notebook文件的精简版本：G:\Project\JupyterLabDir\WindPy_v20230420.ipynb
# 实现的功能包含：
# 1、如何调用Wind API 插件(python)
# 2、参考中金C-REITs指数编制说明（微信收藏的PDF文件），构造C-REITs的价格指数和总回报指数。
# 3、seaborn画相关性矩阵热力图
# 4、波动率计算(日收益率序列如何转化成年化波动率，月度波动率。pandas的方差/标准差函数df.var()/df.std())
# 5、月度涨跌幅/月均收益率的计算
# 6、滚动相关性(pandas.DataFrame.rolling)
# 7、最大回撤率的计算

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from matplotlib.ticker import FuncFormatter
from datetime import datetime
# from adjustText import adjust_text
from WindPy import w

plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显正负号


# 画图坐标轴数字显示为百分比格式
def to_percent(value, position):
    return '{:.2%}'.format(value)  # 显示2位百分数


# plt.style.use('ggplot') # 重启IDE即可复原

#######################################################################################################################
# 1、调用Wind API 插件(python)
#######################################################################################################################
w.start()  # 默认命令超时时间为120秒，如需设置超时时间可以加入waitTime参数，例如waitTime=60,即设置命令超时时间为60秒
w.isconnected()  # 判断WindPy是否已经登录成功
# 获取日时间序列函数WSD
# w.wsd（codes, fields, beginTime, endTime, options）
# 支持股票、债券、基金、期货、指数等多种证券的基本资料、股东信息、市场行情、证券分析、预测评级、财务数据等各种数据。wsd可以支持取 多品种单指标 或者 单品种多指标 的时间序列数据。
# 获取报表数据函数WSET
# w.wset(tableName, options)
# 用来获取数据集信息，包括板块成分、指数成分、ETF申赎成分信息、分级基金明细、融资标的、融券标的、融资融券担保品、回购担保品、 停牌股票、复牌股票、分红送转等报表数据。
# wsd, wss指定多个参数之间用分号;隔开
# 默认不复权 priceAdj=U
# 前复权 priceAdj=F
# 后复权 priceAdj=B
# 默认指定日周期period=D 月周期period=M etc
# ***************************************************************************************************


start_date = "2021-06-21"  # 起始日
end_date = "2023-07-31"  # 截止日
# path = r"D:\Work\DailyReport"  # 保存文件及图片的路径
output_file = f'REITs_Index_Report_{end_date}.xlsx'

# 获取最新的全部REITs成分股
# 板块ID的代码可以通过WIND=》量化=》数据接口=》代码生成器=》数据集WSET=》板块成分=》参数名称SECTORID，编辑=》参数值修改下拉菜单。
# 选择对应的板块比如沪深REITS后自动生成的，这样可以获得想要板块的成分明细。
# 注意：A股的成分股随着参数日期会有变化，但是REITS貌似是最新的27个成分股(as of 20230420)，即使传入的日期参数较早！比如"date=2021-07-25;sectorid=1000041324000000"
reits_list = w.wset("sectorconstituent", "sectorid=1000041324000000", usedf=True)[1]['wind_code'].to_list()

# 区分为产权类和经营权类
# 项目属性(fund_reitsrproperty)：产权类/特许经营类；资产类型(fund__reitstype)：园区基础设施/交通基础设施/仓储物流/...等
df = w.wss(reits_list, "fund_exchangeshortname, fund_reitsrproperty, fund__reitstype", usedf=True)[
    1]  # 基金场内简称：fund_exchangeshortname
reits_list_CQ = df[df["fund_reitsrproperty".upper()] == "产权类"].index.to_list()
reits_list_JY = df[df["fund_reitsrproperty".upper()] == "特许经营类"].index.to_list()

# 大类资产的指数：沪深300(000300.SH)，中证1000，科创50，中债综合全价指数(CBA00203.CS)，南华商品指数，均为价格指数，一般时默认不复权的
assets_list_str = '000300.SH;000852.SH;000688.SH;CBA00203.CS;NH0100.NHF'

# 保存到excel
with pd.ExcelWriter(output_file) as writer:  # 首次创建，默认写模式mode='w'
    df.to_excel(writer, sheet_name='C-REITs列表')


#######################################################################################################################
# 2、构造REITs指数
#######################################################################################################################
# 中金C-REITs指数逻辑是选取2021年6月21日为基期100，所以没有包括首批9个上市首日的涨跌幅，用的是自由流通市值加权法
# 引入除数机制来平滑公司行为造成的市值变动（如增发扩募）与成分股变化（如新增、替换、剔除等）对指数的影响，保证指数的连续性！


def Create_Index(codes_list, output_name):
    # ***************************************************************************************************
    # 数据清洗与准备
    # *******自由流通市值******* #
    # 自由流通市值：实际用的WIND字段为：REITs流通市值(val_mvc)
    # 已经验证: REITs流通市值(val_mvc)= 收盘价(不复权)*REITs场内流通份额(unit_reitsfloortrading)
    df = w.wsd(codes_list, "val_mvc", start_date, end_date, usedf=True)[1]
    df.loc[:, "总自由流通市值"] = df.sum(axis=1)  # 当天所有自由流通市值之和
    df.loc[:, "价格指数"] = 100.00  # 全部初始化100，float类型
    df.loc[:, "指数除数"] = 1.00  # Index Divisor：指数除数，float类型

    # *******自由流通份额******* #
    # 自由流通份额：实际用的WIND字段为：REITs场内流通份额(unit_reitsfloortrading)
    df_shareCnt = w.wsd(codes_list, "unit_reitsfloortrading", start_date, end_date, usedf=True)[1]
    # ***************************************************************************************************
    # 避免BUG：将上市首日之前的份额全部设置为""NaN"(详细参考前2个cell内容)
    # WIND字段：REITs场内流通份额(unit_reitsfloortrading)在实际上市的前几天也有份额。。并非NA。从而导致除数机制失效，后续计算指数时虚高！
    # 借助"close"字段，上市后才有数字，上市之前为NaN；如果close is null，则对应设置份额字段也为None
    df_mask = w.wsd(codes_list, "close", start_date, end_date, "priceAdj=U", usedf=True)[1]  # "priceAdj=U"不复权
    df_mask = df_mask / df_mask  # 归一化，除以自己，变成1或者NaN
    df_shareCnt = df_shareCnt * df_mask  # 首次上市前的份额被修正为NaN
    # ***************************************************************************************************
    # 求和
    df2 = df_shareCnt.sum(axis=1).to_frame("总自由流通份额")  # 当天所有自由流通份额之和

    # *******手动计算 (调整前的)自由流通市值******* #
    # 前一日的份额* 后一天的收盘价
    df_close = w.wsd(codes_list, "close", start_date, end_date, "priceAdj=U", usedf=True)[1]
    # 直接用pandas dataframe矩阵相乘，速度也不慢
    df3 = (df_shareCnt.shift(1)) * (
        df_close
    )
    # 求和
    df3 = df3.sum(axis=1).to_frame('总自由流通市值_调整前')  # 调整前(不含新成分股或者说新的份额)，当天所有自由流通市值之和

    # ***************************************************************************************************
    # 构造REITs指数-价格指数
    # *******3个表合并merge******* #
    df_index = pd.merge(
        pd.merge(
            df[["价格指数", "指数除数", "总自由流通市值"]],
            df2,
            left_index=True, right_index=True
        ),
        df3,
        left_index=True, right_index=True

    )

    # 准备好np arrays，后续进行修改
    index_divisor = df_index['指数除数'].values
    price_index = df_index['价格指数'].values
    mv = df_index['总自由流通市值'].values
    mv_beforeChg = df_index['总自由流通市值_调整前'].values

    # 初始化基期指数和基期除数
    price_index[0] = 100  # 这里将2021-06-21设置为基期100，可能无法反应上市首日的涨跌幅情况。
    index_divisor[0] = df_index['总自由流通市值'].values[0] / 100  # 基期除数：93451865.64222999

    # ***************************************************************************************************
    # 核心逻辑！
    # 循环遍历，计算每期的指数及指数除数
    for i in range(1, len(index_divisor)):  # 跳过基期，因为已经处理过

        index_divisor[i] = index_divisor[i - 1]  # 调整前的指数除数，取上一期值
        price_index[i] = mv_beforeChg[i] / index_divisor[i]  # t期指数(不含新成分股D)=t期自由流通市值(不含新成分股)/调整前的t期指数除数

        price_index_beforeChg = mv[i] / index_divisor[i]  # 调整前t期指数(纳入新成分股D)=t期自由流通市值(含新成分股)/调整前的t期指数除数
        # 调整t期指数除数，使得 调整后的t期指数(纳入新成分股D)=t期指数(不含新成分股D)
        index_divisor[i] = index_divisor[i] * (price_index_beforeChg / price_index[i])

        assert np.allclose(mv[i] / index_divisor[i], mv_beforeChg[i] / index_divisor[i - 1])  # 检查是否满足逻辑，有小数点误差无法完全相等

    # ***************************************************************************************************
    # 更新指数及指数除数
    df_index.loc[:, "指数除数"] = index_divisor
    df_index.loc[:, "价格指数"] = price_index

    # ***************************************************************************************************
    # 构造REITs指数-分红研究
    # 单位累计分红(div_accumulatedperunit)：自基金成立至指定交易日单位分红之和
    df_div = w.wsd(codes_list, "div_accumulatedperunit", start_date, end_date, usedf=True)[1]
    # 将累计分红转化为当期分红
    df_div.fillna(0, inplace=True)  # 将NA填充为0，因为初次分红前默认是NA，如果不修改为0，分红数字和NA相减返回NA，无法得到正确结果
    df_div = df_div - df_div.shift(1)
    df_div.fillna(0, inplace=True)  # 下移一行会导致首行是NA，这里不处理首行应该也没关系，不影响结果
    # 将日期index转换object为日期类型，方便后续使用日期函数相关的操作
    df_div.index = pd.to_datetime(df_div.index, format='%Y-%m-%d')
    # 画图并保存图片
    df_div.plot(figsize=(8, 4))
    # 设置legend在图像外：http://www.kaotop.com/it/19432.html
    # plt.legend(loc='upper left',fontsize=11)
    # plt.legend(bbox_to_anchor=(1.02, 0), loc=3, borderaxespad=0,fontsize=11) # legend在图外右下
    # plt.legend(bbox_to_anchor=(1.02, 1), loc=2, borderaxespad=0,fontsize=11) # legend在图外右上
    plt.legend(bbox_to_anchor=(0.5, -0.8), loc=8, borderaxespad=0, fontsize=11, ncol=4, frameon=False)  # legend在图外正下方,
    # 其中ncol=4表示一行显示几个图例;frameon：bool值，是否绘制图例的外边框，默认值：True
    plt.savefig(f"{output_name}分红.png", dpi=500, bbox_inches='tight')  # 保存成png的格式，dpi取500，300貌似都行，越高的分辨率越高，对应的文件大小越大
    plt.close()  # 为了防止后续plt.savefig()会包含之前的图片产生重叠，关闭掉。也可以使用plt.show()

    # ***************************************************************************************************
    # 构造REITs指数-总回报指数
    df_index.loc[:, "总回报指数"] = 100.00  # 新增字段
    df_index.loc[:, "当期现金分配指数"] = 0.00
    df_index.loc[:, "当期总回报率"] = 1.00
    # 调整列的顺序
    df_index = df_index[["总自由流通市值", "总自由流通份额", "总自由流通市值_调整前", "指数除数", "当期现金分配指数", "当期总回报率", "价格指数", "总回报指数"]]

    # ***************************************************************************************************
    # 构造总回报指数 (这里部分计算可以直接使用向量化的操作，相比迭代效率更高)
    # 第t期现金分配指数=sum(每基金单位现金派息*自由流通基金单位数)/指数除数
    df_index.loc[:, "当期现金分配指数"] = (df_div * df_shareCnt).sum(axis=1) / df_index['指数除数']
    # 第t期总回报率=[(第t期价格指数+第t期现金分派指数)/第t-1期价格指数]-1
    df_index.loc[:, "当期总回报率"] = (df_index["价格指数"] + df_index["当期现金分配指数"]) / df_index["价格指数"].shift(1) - 1
    # 上面shift(1)会产生首行“当期总回报率”的NA值，不过不影响结果。

    # 准备好np arrays，后续进行修改
    return_index = df_index['总回报指数'].values
    return_rate = df_index['当期总回报率'].values

    # 这里不可避免还是要循环
    for i in range(1, len(return_index)):  # 跳过基期，因为已经处理过
        # 第t期总回报指数=第t-1期总回报指数*(1+第t期总回报率)
        return_index[i] = return_index[i - 1] * (1 + return_rate[i])

    # ***************************************************************************************************
    # 更新总回报指数
    df_index.loc[:, "总回报指数"] = return_index
    # 将日期index转换object为日期类型，否则无法用字符串日期进行判断：df_test.index>'2023-03-25'，而且index作为横坐标显示可能有问题
    df_index.index = pd.to_datetime(df_index.index, format='%Y-%m-%d')

    # 保存图片
    df_index[["价格指数", "总回报指数"]].plot(figsize=(8, 4))
    plt.savefig(f"{output_name}.png", dpi=500)  # 保存成png的格式，dpi取500，300貌似都行，越高的分辨率越高，对应的文件大小越大
    plt.close()  # 为了防止后续plt.savefig()会包含之前的图片产生重叠，关闭掉。也可以使用plt.show()

    # 保存到excel
    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
        df_index.to_excel(writer, sheet_name=f'{output_name}')

    return df_index


# ***************************************************************************************************
# 导出到EXCEL多个sheet案例
# # 导出到EXCEL多个sheet需要声明writer对象
# # If you wish to write to more than one sheet in the workbook, it is necessary to specify an ExcelWriter object:
# # Note that creating an ExcelWriter object with a file name that already exists will result in the contents of the
# # existing file being erased.
# with pd.ExcelWriter(path + '\\' +'output.xlsx') as writer:
#     df_div.to_excel(writer, sheet_name='Sheet_name_1')
#     df_index.to_excel(writer, sheet_name='Sheet_name_2')

# # 在原有的EXCEL追加1个sheet:追加模式
# # ExcelWriter can also be used to append to an existing Excel file:
# # 我们应该指定引擎为 openpyxl，而不是默认的 xlsxwriter；否则，我们会得到 xlswriter 不支持 append 模式的错误信息。
# # mode='a',engine='openpyxl'，其中mode的默认值是w，如果ExcelWriter object的文件名是同一个，会覆盖原来的内容。
# with pd.ExcelWriter(path + '\\' + 'output.xlsx', mode='a', engine='openpyxl') as writer:
#     df_index.to_excel(writer, sheet_name='Sheet_name_3')
# ***************************************************************************************************


# ***************************************************************************************************
# 分别构建产权类指数、经营权类指数和总指数，并合并到一张表
df_index1 = Create_Index(reits_list, "C-REITs指数")
df_index2 = Create_Index(reits_list_CQ, "产权类REITs指数")
df_index3 = Create_Index(reits_list_JY, "经营类REITs指数")

df_index = pd.merge(
    pd.merge(
        df_index1[["价格指数", "总回报指数"]].rename(columns={"价格指数": "C-REITs价格指数", "总回报指数": "总回报指数"}),
        df_index2[["价格指数", "总回报指数"]].rename(columns={"价格指数": "产权类REITs价格指数", "总回报指数": "产权类REITs总回报指数"}),
        left_index=True, right_index=True
    ),
    df_index3[["价格指数", "总回报指数"]].rename(columns={"价格指数": "经营类REITs价格指数", "总回报指数": "经营类REITs总回报指数"}),
    left_index=True, right_index=True

)
# 保存图片
df_index[["C-REITs价格指数", "总回报指数", "产权类REITs价格指数", "产权类REITs总回报指数", "经营类REITs价格指数", "经营类REITs总回报指数"]].plot(
    figsize=(8, 4))
plt.savefig("指数对比图.png", dpi=500)  # 保存成png的格式，dpi取500，300貌似都行，越高的分辨率越高，对应的文件大小越大
plt.close()  # 为了防止后续plt.savefig()会包含之前的图片产生重叠，关闭掉。也可以使用plt.show()
#######################################################################################################################


#######################################################################################################################
# 3、相关性矩阵热力图：大类资产
#######################################################################################################################
# 股债指数：分别选取的沪深300(000300.SH)，中证1000，科创50，中债综合全价指数(CBA00203.CS)，南华商品指数，均为价格指数
df_assets = w.wsd(assets_list_str,
                  "pct_chg",  # pct_chg默认是复权，但是如果是价格指数，一般是不复权的
                  start_date, end_date, "priceAdj=U", usedf=True)[1]  # "priceAdj=U"不起作用，因为价格指数一般都是不复权的

# 可选：重命名代码为名称(sec_name)
# DataFrame转字典，df.to_dict()：返回的是复合(嵌套)字典，每列都会单独生成1个字典，以df的index为key，某列为value
# https://www.zhihu.com/question/383246267/answer/2706778845?utm_id=0
df_assets = df_assets.rename(columns=
                             w.wss(assets_list_str,
                                   "sec_name", usedf=True)[1].to_dict()["sec_name".upper()]  # WIND的列默认是大写的字段名称
                             )
# 注意：df_assets原来的列名均为代码'000300.SH;000852.SH;000688.SH;CBA00203.CS;NH0100.NHF'，重命名后修改成下面的列名：
# 对应的sec_name分别是"沪深300","中证1000","科创50","中债-综合全价(总值)指数","南华商品指数"

# 计算REIT指数的日涨跌幅，并合并股债
df_tmp = pd.merge(
    pd.merge(
        # 价格指数的涨跌幅,首行NA也不影响
        ((df_index['C-REITs价格指数'] - df_index['C-REITs价格指数'].shift(1)) / df_index['C-REITs价格指数'].shift(1)).to_frame(
            'REITs价格指数'),
        ((df_index['产权类REITs价格指数'] - df_index['产权类REITs价格指数'].shift(1)) / df_index['产权类REITs价格指数'].shift(1)).to_frame(
            '产权类REITs价格指数'),
        left_index=True, right_index=True
    ),
    ((df_index['经营类REITs价格指数'] - df_index['经营类REITs价格指数'].shift(1)) / df_index['经营类REITs价格指数'].shift(1)).to_frame(
        '经营类REITs价格指数'),
    left_index=True, right_index=True

)

# 2023-08-29 Bo：
df_tmp = df_tmp * 100
# 2023-08-29 Bo:注意上面的涨跌幅乘以100，与WIND涨跌幅口径保持一致，后面统一再除100

df_assets = pd.merge(
    df_tmp,
    df_assets,
    left_index=True, right_index=True
)

# 将日期index转换object为日期类型，方便后续使用日期函数相关的操作
df_assets.index = pd.to_datetime(df_assets.index, format='%Y-%m-%d')

# WIND涨跌幅(pct_chg)简单收益率转换成对数收益率：方便后面相加(注意，对数收益率和简单收益率的相关性结果其实差不多)
# 1+简单收益率=收盘价/前个收盘价=exp(对数收益率)
# 对数收益率=Ln(1+简单收益率)
df_assets = df_assets.apply(lambda x: np.log(1 + x / 100))  # WIND涨跌幅返回值需要除100

# 相关性矩阵热力图
plt.figure(figsize=(6, 3))
sns.heatmap(df_assets.corr(), cmap='RdYlGn', vmax=1, vmin=-1, center=0, annot=True, fmt=".2f")
# http://seaborn.pydata.org/generated/seaborn.heatmap.html?highlight=heatmap#seaborn.heatmap
# sns.heatmap(data,vmax, vmin, center, cmap='rainbow', annot=True)
# annot默认为False，当annot为True时，在heatmap中每个方格写入数据
# center:将数据设置为图例中的均值数据，即图例中心的数据值；通过设置center值，可以调整生成的图像颜色的整体深浅；
# vmax=1, vmin=-1，图例的上限和下限
# fmt: String formatting code to use when adding annotations.

# 保存图片
plt.savefig("大类资产相关性.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df_assets.corr().to_excel(writer, sheet_name='大类资产相关性')
# ***************************************************************************************************


# ***************************************************************************************************
# 相关性矩阵热力图：各个REITs之间
# WIND涨跌幅pct_chg经过验证用的复权价格计算的涨跌幅pct_chg。
df = w.wsd(reits_list, "pct_chg", start_date, end_date, usedf=True)[1]  # pct_chg默认是复权
# WIND涨跌幅(pct_chg)简单收益率转换成对数收益率
df = df.apply(lambda x: np.log(1 + x / 100))  # WIND涨跌幅返回值需要除100

# 可选：重命名列：代码=》简称
df = df.rename(columns=w.wss(reits_list, "fund_exchangeshortname",  # 基金场内简称：fund_exchangeshortname
                             usedf=True)[1].to_dict()["fund_exchangeshortname".upper()])
# DataFrame转字典，df.to_dict()：返回的是复合(嵌套)字典，每列都会单独生成1个字典，以df的index为key，某列为value

plt.figure(figsize=(15, 15))
sns.heatmap(df.corr(), cmap='RdYlGn', vmax=1, vmin=-1, center=0, annot=True, fmt=".2f")

# 保存图片
plt.savefig("REITs相关性.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df.corr().to_excel(writer, sheet_name='REITs相关性')
#######################################################################################################################


#######################################################################################################################
# 4、波动率计算
#######################################################################################################################
# 如果一年以250交易日计算，即50周 年化标准差=日收益标准差sqrt(250)=周收益标准差sqrt(50)
# 年化波动率=收益标准差*(n^0.5)。其中：计算周期为日，对应n为 250；计算周期为周，对应n为 52；计算周期为月，对应n为 12；计算周期为年，对应n为 1。
# 标准差是衡量风险的常用标准，是与时间期限相关的概念，例如日标准差、周标准差、月标准差、年标准差等等。在风险评价中，常用的是年标准差。
# 但是不同时间期限的标准差不能直接比较，例如，日标准差1%和月标准差5%。因此，要把不同期限的标准差都转化为统一的以年为期限，才能进行比较。
# 日标准差转化为年标准差的公式为，年标准差=日标准差*(365)^0.5。
# 一般地，用日收益率的时间序列计算的标准差是日收益率分布的标准差，转化为年化的标准差通常采用时间平方根公式，即：年化标准差=sqrt(T)*日收益率的标准差，
# 其中sqrt()指平方根函数，股票市场通常取每年有T=252个交易日。类似地，如果是日收益率的标准差想转化成月度的标准差，取每月有T=22个交易日。

# pandas-var样本方差、std样本标准差:
# df.var()  # 样本方差
# df.var(ddof=0)  # 总体方差，To have the same behaviour as numpy.var, use ddof=0 (instead of the default ddof=1)
# df.std()  # 样本标准差
# df.std(ddof=0)  # 总体标准差，To have the same behaviour as numpy.std, use ddof=0 (instead of the default ddof=1)

# https://www.malaoshi.top/show_1IX1EVy4DPwc.html
# http://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.var.html?highlight=var#pandas.DataFrame.var
# http://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.std.html?highlight=std#pandas.DataFrame.std
# 注意：df.std()样本标准差等价于excel的STDEV.S

# ***************************************************************************************************
# 计算每个REITs的月度波动率
# 使用pct_chg获取涨跌幅序列,包含上市首日的涨跌幅(复权)的波动率！
df_risk = w.wsd(reits_list, "pct_chg", start_date, end_date, "priceAdj=U", usedf=True)[1]  # "priceAdj=U"不起作用，实际用的复权
# 简单收益率转换成对数收益率
df_risk = df_risk.apply(lambda x: np.log(1 + x / 100))  # WIND涨跌幅返回值需要除100
# df_risk.std()*250**0.5  # 年化波动率
df_risk = (df_risk.std() * (250 / 12) ** 0.5).to_frame("月度波动率")  # 月度波动率
#######################################################################################################################


#######################################################################################################################
# 5、月均收益率
#######################################################################################################################
# 方法一(有BUG)，直接使用WIND函数，月度涨跌幅pct_chg：传入参数Period=M，默认是D
# Period取值周期。参数值含义如下：
# D：天，默认值
# W：周，
# M：月，
# Q：季度，
# S：半年，
# Y：年
# 注意：这里WIND pct_chg计算首月涨跌幅时有BUG！！！非上市首月的涨跌幅则没问题。
# pct_chg计算首月涨跌幅的分母似乎是上市首日的收盘价，而不是发行价格；而pct_chg默认的日涨跌幅序列，首日的涨跌幅分母是发行价，没问题。

# 方法二(推荐！)，基于日涨跌幅，手动分组，计算月涨跌幅！
# Pandas按月份和年份分组（日期为datetime64[ns]），并汇总
# https://www.cnpython.com/qa/1311357
# 计算每个REITs的月均收益率
df_test = w.wsd(reits_list, "pct_chg", start_date, end_date, "priceAdj=U", usedf=True)[1]  # "priceAdj=U"不起作用，实际用的复权

# WIND涨跌幅(pct_chg)简单收益率转换成对数收益率，因为对数收益率方便直接相加！
df_test = df_test.apply(lambda x: np.log(1 + x / 100))  # WIND涨跌幅返回值需要除100
# 将日期index转换object为日期类型，方便后续使用日期函数相关的操作
df_test.index = pd.to_datetime(df_test.index, format='%Y-%m-%d')

# 按月分组并求和
# 方法一：利用日期date中获取dt.year和dt.month
# 注意，这里的日期刚好在index中，可以直接df.index.year。否则如果日期是某一列中，则应该接一个dt比如：df['date'].dt.year
# df_test = df_test.groupby([df_test.index.year.rename('year'),df_test.index.month.rename('month')]).sum()  # 对数收益率可以直接相加

# 方法二(推荐!)：datetime还支持to_period转换，因此我们可以按月对所有内容进行分组
# 注意，这里的日期刚好在index中，可以直接df.index.to_period('M')。否则如果日期是某一列中，则应该接一个dt比如：df['date'].dt.to_period('M')
df_test = df_test.groupby(df_test.index.to_period('M')).sum()  # 这里sum()会将NA当成0处理，sum NA结果是0

# 对数收益率转简单收益率，即月涨跌幅
# 1+简单收益率=收盘价/前个收盘价=exp(对数收益率)
# 简单收益率=exp(对数收益率)-1
df_test = df_test.apply(lambda x: np.exp(x) - 1)

# ***************************************************************************************************
# 注意！！！这里必须修正上市以来的月份数，计算月均收益率
# 上面groupby分组后再sum，会将NA值变成0，再计算月均收益率时，月份会多算，这里需要将上市首月前的0，修正为NA，方便计算平均值！
df_mask = w.wsd(reits_list, "pct_chg", start_date, end_date, "priceAdj=U", usedf=True)[1]  # "priceAdj=U"不起作用，实际用的复权
# WIND涨跌幅(pct_chg)简单收益率转换成对数收益率，因为对数收益率方便直接相加！
df_mask = df_mask.apply(lambda x: np.log(1 + x / 100))  # WIND涨跌幅返回值需要除100
# 将日期index转换object为日期类型，方便后续使用日期函数相关的操作
df_mask.index = pd.to_datetime(df_mask.index, format='%Y-%m-%d')
# 分组并判断组内是否为NA
df_mask = df_mask.groupby(df_mask.index.to_period('M')).apply(lambda x: x.isnull().all())  # 组内的全部为NA

# pandas.DataFrame.applymap
# Apply a function to a Dataframe elementwise.
# This method applies a function that accepts and returns a scalar to every element of a DataFrame.
# 对整个数据表作用时需要改成df.applymap()函数，df.apply()是应用于pd.Series或列或行，这里如果用df.apply()会报错
# https://blog.csdn.net/weixin_55674264/article/details/122592950
df_mask = df_mask.applymap(lambda x: None if x else 1.0)  # True(即组内全部为NA)则值修正为NA，否则设置为1
# 修正后的月均涨跌幅/收益率：(df_mask*df_test).mean()
# ***************************************************************************************************

# 合并修正后的月均收益率，之前的月度波动率，再合并基金场内简称：fund_exchangeshortname
df_risk = pd.merge(
    pd.merge(
        df_risk,
        (df_mask * df_test).mean().to_frame('月均收益率'),  # 修正后的月均涨跌幅/收益率
        left_index=True, right_index=True
    ),
    w.wss(reits_list, "fund_exchangeshortname", usedf=True)[1].rename(columns={"fund_exchangeshortname".upper(): "简称"}),
    left_index=True, right_index=True
)

# 画每个REITs风险收益的散点图
plt.figure(figsize=(8, 4))
sns.scatterplot(data=df_risk, x='月度波动率', y='月均收益率', hue="简称")  # hue也可以设置成代码，即hue=df_risk.index
plt.gca().xaxis.set_major_formatter(FuncFormatter(to_percent))  # x轴百分位格式
plt.gca().yaxis.set_major_formatter(FuncFormatter(to_percent))  # y轴百分位格式
# 设置legend在图像外：http://www.kaotop.com/it/19432.html
# plt.legend(loc='upper left',fontsize=11)
# plt.legend(bbox_to_anchor=(1.02, 0), loc=3, borderaxespad=0,fontsize=11) # legend在图外右下
# plt.legend(bbox_to_anchor=(1.02, 1), loc=2, borderaxespad=0,fontsize=11) # legend在图外右上
plt.legend(bbox_to_anchor=(0.5, -0.7), loc=8, borderaxespad=0, fontsize=11, ncol=4, frameon=False)  # legend在图外正下方,
# 其中ncol=4表示一行显示几个图例;frameon：bool值，是否绘制图例的外边框，默认值：True
# 给散点加标签
# texts = [plt.gca().text(df_risk['月度波动率'].iloc[i], df_risk['月均收益率'].iloc[i],
#                         df_risk['简称'].iloc[i]) for i in range(len(df_risk))]  # text(x坐标,y坐标,文本内容)
# 保存图片
plt.savefig("REITs风险收益.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df_risk.to_excel(writer, sheet_name='REITs风险收益')

# ***************************************************************************************************
# 同上：计算大类资产的风险收益
# 波动率：前面已经保存并计算大类资产的每日对数收益率序列df_assets
df_risk = (df_assets.std() * (250 / 12) ** 0.5).to_frame("月度波动率")  # 月度波动率

# 按月分组计算月均收益率(这里不需要修正月份数，因为区间起始日2021-06-21，全部大类资产都有数据)
df_test = df_assets.groupby(df_assets.index.to_period('M')).sum()
# 对数收益率转简单收益率，即月涨跌幅
# 1+简单收益率=收盘价/前个收盘价=exp(对数收益率)
# 简单收益率=exp(对数收益率)-1
df_test = df_test.apply(lambda x: np.exp(x) - 1)

# 合并
df_risk = pd.merge(
    df_risk,
    df_test.mean().to_frame('月均收益率'),
    left_index=True, right_index=True
)

# 画大类资产风险收益的散点图
# plt.figure(figsize=(8, 4))
sns.scatterplot(data=df_risk, x='月度波动率', y='月均收益率', hue=df_risk.index)
plt.gca().xaxis.set_major_formatter(FuncFormatter(to_percent))  # x轴百分位格式
plt.gca().yaxis.set_major_formatter(FuncFormatter(to_percent))  # y轴百分位格式
plt.legend(bbox_to_anchor=(0.5, -0.3), loc=8, borderaxespad=0, fontsize=11, ncol=3, frameon=False)  # legend在图外正下方
# 给散点加标签
texts = [plt.gca().text(df_risk['月度波动率'].iloc[i], df_risk['月均收益率'].iloc[i],
                        df_risk.index[i]) for i in range(len(df_risk))]  # text(x坐标,y坐标,文本内容)
# 保存图片
plt.savefig("大类资产风险收益.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df_risk.to_excel(writer, sheet_name='大类资产风险收益')
#######################################################################################################################


#######################################################################################################################
# 6、滚动相关性
#######################################################################################################################
# pandas.DataFrame.rolling
# http://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.rolling.html?highlight=rolling#pandas.DataFrame.rolling
# DataFrame.rolling(window, min_periods=None, center=False, win_type=None, on=None, axis=0, closed=None, step=None, method='single')
# 参数window说明:
# int, timedelta, str, offset, or BaseIndexer subclass
# Size of the moving window.
# If an integer, the fixed number of observations used for each window.
# If a timedelta, str, or offset, the time period of each window.
# Each window will be a variable sized based on the observations included in the time-period.
# This is only valid for datetimelike indexes.

# 大类资产的滚动相关性计算(经过excel验证，结果正确！)，前面已经保存并计算大类资产的每日对数收益率序列df_assets
df_assets.rolling(30).corr()  # 注意，这里返回值的每一行(即每一天的近30天滚动相关性)都是一个相关性矩阵

# MultiIndex如何进行索引操作，方法一[('第1级名称','第2级名称',...,'第n级名称')],方法二['第1级名称']['第2级名称']...['第n级名称']
# 索引案例！
# df_test.rolling(30).corr().unstack()[('000300.SH','CBA00203.CS')]
# df_test.rolling(30).corr().unstack()['000300.SH','CBA00203.CS']
# df_test.rolling(30).corr().unstack()['000300.SH']['CBA00203.CS']


# 取对应的列出来：可以查看当前的列名：print(df_assets.rolling(30).corr().unstack().columns)，前面已经替换代码为sec_name了
df_test = df_assets.rolling(30).corr().unstack()[
    [('REITs价格指数', '中债-综合全价(总值)指数'), ('REITs价格指数', '沪深300')]]  # unstack将多级的index放到column维度。

# pandas如何重命名MultiIndex列
# https://www.cnblogs.com/cgmcoding/p/14168268.html
df_test.columns = df_test.columns.to_flat_index()  # 不能少！！to_flat_index去掉多层次，否则后续重命名列不起作用。
# 列重命名
df_test = df_test.rename(columns={('REITs价格指数', '中债-综合全价(总值)指数'): "REITs-债券", ('REITs价格指数', '沪深300'): "REITs-股票"})

# 画图，去掉上市前2个月的数据，从21年8月开始
df_test.plot(xlim=('2021-08-01', end_date), ylim=(-1, 1), figsize=(10, 5))  # df.plot可以直接传入figsize, xlim, ylim等参数
# 保存图片
plt.savefig("滚动相关性.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df_test.to_excel(writer, sheet_name='滚动相关性')
#######################################################################################################################


#######################################################################################################################
# 7、最大回撤
#######################################################################################################################
# 方法一：使用WIND函数：risk_maxdownside 并传入参数startDate=2022-11-15;endDate=2022-12-28
# 貌似现在跟方法二计算的结果有些许差别，而且可能也有前面类似的问题，即上市首日的计算可能不对。
# wind计算的全部REITs在2022年的最大回撤
w.wss(reits_list,
      "risk_maxdownside, fund_exchangeshortname",
      "startDate=2022-01-01;endDate=2022-12-31", usedf=True)[1].sort_values(by='risk_maxdownside'.upper())

# 方法二：
# 使用自定义最大回撤函数：TBD：目前也没有考虑上市首日的回撤情况(相比于发行价格)
df_tmp = w.wsd(reits_list,
               "close",
               start_date, end_date, "priceAdj=B", usedf=True)[1]  # "priceAdj=B"后复权，用来计算最大回撤
# 可选：重命名代码为名称
df_tmp = df_tmp.rename(columns=w.wss(reits_list, "fund_exchangeshortname",  # 基金场内简称：fund_exchangeshortname
                                     usedf=True)[1].to_dict()["fund_exchangeshortname".upper()])


# 最大回撤算法：对于某区间的每一个价格(或净值)，分别回溯找到其对应的最大回撤值，然后再找出其中最大的回撤值
# 最大回撤函数：传入一组价格(或净值)数据，返回最大回撤值
def MaxDrawDown(prices):  # 可接受list, np.array, pandas.series

    # 移除NA值，因为WIND函数返回的序列中可能包含NA值。否则后续因为存在NA值而无法计算最大回撤
    prices = np.array(prices)  # 统一转换成numpy.arrays，不管之前是List还是pandas.series
    prices = prices[np.where(~np.isnan(prices))]  # np.isnan判断是否为NaN值，np.where(condition)返回的是坐标
    max_draw_down = 0.0
    for i in range(1, len(prices)):  # 跳过第1个，因为第1个数没有回撤概念，从第2个开始才有回撤
        previous_max = max(prices[:i])  # 对于每一个价格，回溯其之前的最大值
        draw_down = (prices[i] - previous_max) / previous_max  # 计算回撤率
        # 如果某个回撤值更低，则替换最大回撤
        if draw_down < max_draw_down:
            max_draw_down = draw_down

    return max_draw_down


# 测试
# MaxDrawDown(df_tmp['180401.SZ'])
# MaxDrawDown([1,0.9, 0.8,0.9,1,0.9])
# MaxDrawDown(df_tmp['180401.SZ'].values)

# 使用apply函数计算每列的最大回撤值
df_tmp = df_tmp.apply(MaxDrawDown, axis=0).to_frame('最大回撤率').sort_values(by='最大回撤率')  # apply函数默认axis=0

sns.barplot(data=df_tmp, y=df_tmp.index, x="最大回撤率")
plt.gca().xaxis.set_major_formatter(FuncFormatter(to_percent))  # x轴百分位格式
# 保存图片
plt.savefig("最大回撤率.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠
# 保存到excel
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:  # 追加模式并指定引擎
    df_tmp.to_excel(writer, sheet_name='最大回撤率')


# ***************************************************************************************************
# 获取某个标的的最大回撤明细信息，并画图！
# 最大回撤函数：传入一组价格(或净值)数据，返回最大回撤值及相关明细信息
def MaxDrawDown2(prices):  # 可接受list, np.array, pandas.series

    index_offset = 0  # 用于记录原始数据(可能包含NA)的index位置
    max_val = 0
    max_index = 0
    min_val = 0
    min_index = 0
    max_draw_down = 0.0

    # 移除NA值，因为WIND函数返回的序列中可能包含NA值。否则后续因为存在NA值而无法计算最大回撤
    prices = np.array(prices)  # 统一转换成numpy.arrays，不管之前是List还是pandas.series
    prices = np.squeeze(prices)  # 如果是2维则转化成1维，防止意外BUG
    # numpy.squeeze(a, axis=None)
    # Remove axes of length one from a.
    # https://numpy.org/doc/stable/reference/generated/numpy.squeeze.html#numpy-squeeze

    # 记录首个非NA值的位置
    index_offset = np.where(~np.isnan(prices))[0][0]  # 第1个[0]是因为貌似这里np.where返回的是tuple，第2个[0]才是取第1个元素位置
    # 例如：(array([182, 183, 184, 185, 186, 187, 188], dtype=int64),)

    # 移除NA值
    prices = prices[np.where(~np.isnan(prices))]  # np.isnan判断是否为NaN值，np.where(condition)返回的是坐标

    for i in range(1, len(prices)):  # 跳过第1个，因为第1个数没有回撤概念，从第2个开始才有回撤
        previous_max = max(prices[:i])  # 对于每一个价格，回溯其之前的最大值
        draw_down = (prices[i] - previous_max) / previous_max  # 计算回撤率
        # 如果某个回撤值更低，则替换最大回撤
        if draw_down < max_draw_down:
            max_draw_down = draw_down
            max_val = previous_max
            max_index = np.argmax(prices[:i])
            min_val = prices[i]
            min_index = i

    return (max_draw_down, max_index + index_offset, max_val, min_index + index_offset, min_val)


# numpy.argmax(a, axis=None, out=None, *, keepdims=<no value>)
# Returns the indices of the maximum values along an axis.
# https://numpy.org/doc/stable/reference/generated/numpy.argmax.html#numpy-argmax
# 参数axis: int, optional
# By default, the index is into the flattened array, otherwise along the specified axis.

# 画图
code = '180101.SZ'
df_chart = w.wsd(code,
                 "close",
                 start_date, end_date, "priceAdj=B", usedf=True)[1]  # 用复权价格计算最大回撤率
# 获取最大回撤率计算明细结果
(max_draw_down, max_index, max_val, min_index, min_val) = MaxDrawDown2(df_chart)
# 检查结果，df_chart可能是pd.DataFrame或者pd.Series，取坐标的时候有些许不同
print(max_draw_down, max_index, max_val, min_index, min_val)
a = df_chart.iloc[max_index][0] if isinstance(df_chart, pd.DataFrame) else df_chart[max_index]
b = df_chart.iloc[min_index][0] if isinstance(df_chart, pd.DataFrame) else df_chart[min_index]
print(a, max_val)
print(b, min_val)
print((b - a) / a, max_draw_down)
# 将日期index转换object为日期类型，方便后续使用日期函数相关的操作
df_chart.index = pd.to_datetime(df_chart.index, format='%Y-%m-%d')

df_chart.plot()
# 获取最大回撤区间的横坐标，并绘制红色虚线
max_index = df_chart.index[max_index]
min_index = df_chart.index[min_index]
plt.plot([max_index, max_index], [min_val - 0.1, max_val + 0.1], "r--")  # 绘制红色虚线：plt.plot(x, y) 其中x和y是横坐标和纵坐标
plt.plot([min_index, min_index], [min_val - 0.1, max_val + 0.1], "r--")  # 绘制红色虚线：plt.plot(x, y) 其中x和y是横坐标和纵坐标
# 添加文字标注: max_index.date()仅显示日期，不显示小时
plt.gca().text(max_index, max_val + 0.1, f'{max_index.date()},最高值{max_val:.3f}', color='r')  # text(x坐标,y坐标,文本内容)
plt.gca().text(min_index, min_val - 0.1, f'{min_index.date()},最低值{min_val:.3f}', color='r')  # text(x坐标,y坐标,文本内容)
plt.title(f'最大回撤率({code}): {max_draw_down * 100:.2f}%')
plt.legend()
# 保存图片
plt.savefig(f"最大回撤率 {code}.png", dpi=500, bbox_inches='tight')
plt.close()  # 防止后续图片重叠

#######################################################################################################################
