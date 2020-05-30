# 导入通用模块
import numpy as np
import pandas as pd 
import os 

# 读取数据表
sheet = pd.read_excel('需要提取经纬度的表.xlsx')

# 固定文本
## 起始头
str_star = """
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2">
	<Document id="TbuluKmlVersion2">
		<name><![CDATA[
            """
## 段落1 
ph1 = """
]]></name>
		<Style id="TrackStyle_n">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=BYohR0x5OLAr0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LineStyle>
				<color>ee0000ff</color>
				<width>5.000000</width>
			</LineStyle>
		</Style>
		<Style id="TrackStyle_h">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=BYohR0x5OLAr0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LineStyle>
				<color>ff0000ff</color>
				<width>7.000000</width>
			</LineStyle>
		</Style>
		<StyleMap id="TrackStyle">
			<Pair>
				<key>normal</key>
				<styleUrl>#TrackStyle_n</styleUrl>
			</Pair>
			<Pair>
				<key>highlight</key>
				<styleUrl>#TrackStyle_h</styleUrl>
			</Pair>
		</StyleMap>
		<Style id="LineStringStyle">
			<LineStyle>
				<color>dd5cba55</color>
				<width>6.000000</width>
			</LineStyle>
		</Style>
		<Style id="startPointStyle">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=n%2BVE0NaFmzwndk%2BZ75NU9A%3D%3D%0A</href>
				</Icon>
			</IconStyle>
		</Style>
		<Style id="endPointStyle">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=0Q3GJ5o%2FHpsr0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
		</Style>
		<Style id="MarkerStylePicture">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=vMzvc3dTBl0r0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LabelStyle>
				<color>ff00ffff</color>
			</LabelStyle>
		</Style>
		<Style id="MarkerStyleSound">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=poL0RfxWHTIr0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LabelStyle>
				<color>ff00ffff</color>
			</LabelStyle>
		</Style>
		<Style id="MarkerStyleVideo">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=vMzvc3dTBl0r0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LabelStyle>
				<color>ff00ffff</color>
			</LabelStyle>
		</Style>
		<Style id="MarkerStyleText">
			<IconStyle>
				<Icon>
					<href>https://files.2bulu.com/f/d1?downParams=s2An6%2B3%2BIWUr0OBxohfNOA%3D%3D%0A</href>
				</Icon>
			</IconStyle>
			<LabelStyle>
				<color>ff00ffff</color>
			</LabelStyle>
		</Style>
		<ExtendedData>
			<Data name="TrackId">
				<value>11131023</value>
			</Data>
			<Data name="SportTypeId">
				<value>18</value>
			</Data>
			<Data name="BeginTime">
				<value>1568705029538</value>
			</Data>
			<Data name="EndTime">
				<value>1568706086462</value>
			</Data>
			<Data name="Privacy">
				<value>0</value>
			</Data>
			<Data name="PauseTime">
				<value>58459</value>
			</Data>
			<Data name="TrackTypeId">
				<value>8</value>
			</Data>
			<Data name="TrackTags">
				<value><![CDATA[]]></value>
			</Data>
			<Data name="PosStartName">
				<value><![CDATA[]]></value>
			</Data>
			<Data name="PosEndName">
				<value><![CDATA[]]></value>
			</Data>
			<Data name="ProductVersion">
				<value>6.4.60.7</value>
			</Data>
			<Data name="CreaterId">
				<value>22585544</value>
			</Data>
			<Data name="CreaterName">
				<value>zs_4916946</value>
			</Data>
			<Data name="OriginCreaterId">
				<value>0</value>
			</Data>
			<Data name="OriginCreaterName">
			</Data>
			<Data name="OriginCreaterNickname">
			</Data>
			<Data name="Step">
				<value>1597</value>
			</Data>
			<Data name="Calorie">
				<value>70.890923</value>
			</Data>
			<Data name="Mileage">
				<value>1007.251183</value>
			</Data>
			<Data name="Source">
				<value>0</value>
			</Data>
			<Data name="FragmentId">
				<value>0</value>
			</Data>
		</ExtendedData>
		<Folder id="TbuluHisPointFolder">
			<name><![CDATA[标注点]]></name>
			<Placemark id="startPoint">
				<name><![CDATA[起点]]></name>
				<description><![CDATA[<div>经度：
                """
# 段落2
ph2 = """
</div>]]></description>
				<styleUrl>#startPointStyle</styleUrl>
				<Point>
					<coordinates>
                    """

# 段落3
ph3 = """
</coordinates>
				</Point>
			</Placemark>
			<Placemark id="endPoint">
				<name><![CDATA[终点]]></name>
				<description><![CDATA[<div>经度：
                    """

# 段落4
ph4 = """
</div>]]></description>
				<styleUrl>#endPointStyle</styleUrl>
				<Point>
					<coordinates>
                    """

# 段落5
ph5 = """
</coordinates>
				</Point>
			</Placemark>
		</Folder>
		<Folder id="TbuluTrackFolder">
			<name><![CDATA[轨迹]]></name>
			<Placemark>
				<styleUrl>#TrackStyle</styleUrl>
				<gx:Track>
					<ExtendedData>
						<Data name="GxTrackExtendedData">
							<value>1.320000,0.0;1.330000,0.0;1.220000,0.0;1.300000,0.0;1.210000,0.0;1.290000,0.0;1.250000,0.0;1.180000,0.0;1.300000,0.0;1.330000,0.0;1.200000,0.0;0.290000,0.0;0.980000,0.0;1.090000,0.0;1.230000,0.0;1.060000,0.0;1.570000,0.0;1.290000,0.0;1.290000,0.0;1.290000,0.0;1.430000,0.0;1.260000,0.0;1.090000,0.0;1.340000,0.0;0.980000,0.0;0.860000,0.0;1.400000,0.0;1.170000,0.0;1.080000,0.0;1.050000,0.0;1.170000,0.0;1.250000,0.0;1.290000,0.0;1.230000,0.0;0.870000,0.0;1.320000,0.0;1.240000,0.0;1.340000,0.0;1.080000,0.0;1.060000,0.0;1.220000,0.0;1.160000,0.0;1.040000,0.0;1.110000,0.0;1.240000,0.0;1.330000,0.0;1.050000,0.0;1.230000,0.0;1.370000,0.0;0.810000,0.0;1.170000,0.0;1.070000,0.0;1.360000,0.0;0.880000,0.0;0.890000,0.0;1.240000,0.0;0.810000,0.0;1.110000,0.0;1.030000,0.0;0.960000,0.0;1.010000,0.0;0.930000,0.0;0.790000,0.0;1.110000,0.0;0.900000,0.0;0.900000,0.0;1.060000,0.0;0.630000,0.0;1.290000,0.0;1.360000,0.0;1.160000,0.0;1.250000,0.0;1.050000,0.0;1.350000,0.0;1.140000,0.0;0.870000,0.0;0.870000,0.0;1.120000,0.0;0.820000,0.0;0.840000,0.0;1.240000,0.0;1.260000,0.0;1.250000,0.0;0.920000,0.0;1.100000,0.0;0.900000,0.0;1.220000,0.0;1.150000,0.0;1.030000,0.0;1.040000,0.0;0.930000,0.0;1.040000,0.0;0.760000,0.0;1.280000,0.0;1.290000,0.0;</value>
						</Data>
					</ExtendedData>
					<gx:coord>
                    """

# 段落6
ph6 = """
</gx:coord>
					<gx:coord>
                    """

# 段落7
ph7 = """
</gx:coord>
					<when>
                    """

# 段落8
ph8 = """
</when>
					<when>
                    """

str_end = """
</when>
				</gx:Track>
			</Placemark>
		</Folder>
	</Document>
</kml>
"""

# 遍历表格中的数据
for i in range(len(sheet.index)):
    name = sheet.iloc[i][0]
    longitude_star = str(sheet.iloc[i][1])
    latitude_star = str(sheet.iloc[i][2])
    longitude_end = str(sheet.iloc[i][3])
    latitude_end = str(sheet.iloc[i][4])
    altitude = str(sheet.iloc[i][5])
    date = str(sheet.iloc[i][6])
    coordinates_star = longitude_star + "," + latitude_star + "," + altitude
    coordinates_end = longitude_end + "," + latitude_end + "," + altitude
    coord_star = longitude_star + " " + latitude_star + " " + altitude
    coord_end = longitude_end + " " + latitude_end + " " + altitude
    result = str_star + name + ph1 + longitude_star + "</div><div>纬度：" + latitude_star + "</div><div>海拔：" + altitude + "</div><div>时间：" + date + ph2 + coordinates_star + ph3 + longitude_end + "</div><div>纬度：" + latitude_end + "</div><div>海拔：" + "</div><div>时间：" + date + ph4 + coordinates_end + ph5 + coord_star + ph6 + coord_end + ph7 + date + ph8 + date + str_end
    with open('{}.kml'.format(name),'w') as f:
        f.write(result)
        f.close()