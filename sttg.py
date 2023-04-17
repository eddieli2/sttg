#1.分别找出19日和23日每位客户经理名下的自营关系
#2.将找出的自营关系置于与客户经理并列的位置，找出的自营关系的部门还原为和梅提供的部门，主自营客户经理人力ID号还原和梅提供的ID
#3.计算每位客户经理名下所有用户的存款总额
#4.计算23日与19日的差值得出每位客户经理的新增的存款数
#5.将二类卡的开户数与存款金额与客户经理进行对应
#6.根据给定公式计算每个客户经理的增配财务资源
import pandas as pd
df19 = pd.read_excel("C:/Users/Administrator/Desktop/hdsjfx/0919.xlsx")
df23 = pd.read_excel("C:/Users/Administrator/Desktop/hdsjfx/0923.xlsx")
df18 = pd.read_excel("C:/Users/Administrator/Desktop/hdsjfx/sfzh.xlsx")

df19_2 = df19.loc[df19['证件号码'].isin(df18['身份证号']), :]
df23_2 = df23.loc[df23['证件号码'].isin(df18['身份证号']), :]

df19_2 #df19中包含的行内员工,df19中包含的88位行内员工
df23_2 #df23中包含的行内员工，df23中包含的89位行内员工


df19_3 = df19.drop(df19_2.index)
df23_3 = df23.drop(df23_2.index)

df19_3 #从df19中去除他包含的所有行内员工
df23_3 #从df23中去除他包含的所有行内员工

#处理df19_2使员工的人力ID和部门回归到和梅的表
df18 = df18.rename(columns={'身份证号':'证件号码'})
df19_2 = pd.merge(left=df19_2,right=df18,on='证件号码')
df19_2['主自营客户经理人力ID号'] = df19_2['员工 ID']
df19_2['主自营机构'] = df19_2['部门']

#处理df23_2使员工的人力ID和部门回归到和梅的表
df18 = df18.rename(columns={'身份证号':'证件号码'})
df23_2 = pd.merge(left=df23_2,right=df18,on='证件号码')
df23_2['主自营客户经理人力ID号'] = df23_2['员工 ID']
df23_2['主自营机构'] = df23_2['部门']

#将经过处理后的行内员工合并到纯客户表中
df19_4 = pd.concat([df19_3, df19_2], axis=0)
df19_4 #将行内员工合并到df19_3中
df23_4 = pd.concat([df23_3, df23_2], axis=0)
df23_4 #将行内员工合并到df23_3中



df19_5 = df19_4.groupby(['主自营客户经理人力ID号'])['储蓄','个人结构性存款'].sum()
df19_5
df23_5 = df23_4.groupby(['主自营客户经理人力ID号'])['储蓄','个人结构性存款'].sum()
df23_5


df18 = df18.rename(columns={"员工 ID":"主自营客户经理人力ID号"})
df19_5 = pd.merge(left=df19_5,right=df18,on='主自营客户经理人力ID号')
df19_5['主自营客户经理人力ID号'] = df18['姓名']
df23_5 = pd.merge(left=df23_5,right=df18,on='主自营客户经理人力ID号')
df23_5['主自营客户经理人力ID号'] = df18['姓名']


df19_5.columns.insert(2,'存款合计')
df23_5.columns.insert(2,'存款合计')

df19_5['存款合计'] = df19_5['储蓄']+df19_5['个人结构性存款']
df19_5['存款合计']
df23_5['存款合计'] = df23_5['储蓄']+df23_5['个人结构性存款']
df23_5['存款合计'] 

df_final = pd.DataFrame({'客户经理名字':df23_5['主自营客户经理人力ID号'],'19日个人时点存款':df19_5['存款合计'],'23日个人时点存款':df23_5['存款合计'],'新增金额':df23_5['存款合计']-df19_5['存款合计']})

df_final.to_excel("统计结果.xlsx", sheet_name="Output",startrow=1, startcol=1, index=True, header=True,na_rep="<NA>", inf_rep="<INF>")
