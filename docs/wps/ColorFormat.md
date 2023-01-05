**ColorFormat**



代表单色对象的颜色，或以渐变/图案填充的对象的前景色或背景色。

**说明**

使用下表中列出的属性之一返回 [**ColorFormat** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ColorFormat/ColorFormat%20.htm#jsObject_ColorFormat)对象。

| 使用此属性         | 结合此对象       | 返回一个代表以下颜色的 ColorFormat 对象    |
| ------------------ | ---------------- | ------------------------------------------ |
| **BackColor**      | **FillFormat**   | 背景填充颜色（在底纹或带图案的填充中使用） |
| **ForeColor**      | **FillFormat**   | 前景填充颜色（或单色填充的填充颜色）       |
| **BackColor**      | **LineFormat**   | 背景线条颜色（在带图案线条中使用）         |
| **ForeColor**      | **LineFormat**   | 前景线条颜色（或实线的线条颜色）           |
| **ForeColor**      | **ShadowFormat** | 阴影颜色                                   |
| **ExtrusionColor** | **ThreeDFormat** | 突出对象的边的颜色                         |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Brightness**       | 返回一个 **Number** 类型的值，代表指定形状颜色的亮度。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ObjectThemeColor** | 返回或设置一个 **WdThemeColorIndex** 常量，该常量表示颜色格式的主题颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RGB**              | 返回或设置指定颜色的 RGB 值（红-绿-蓝值），**Number** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**             | 返回或设置形状的颜色类型。**MsoColorType** 类型，只读。      |

**成员属性**

#### **ColorFormat.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.Brightness**

返回一个 **Number** 类型的值，代表指定形状颜色的亮度。可读/写。

**语法**

**express.Brightness**

*express*   一个代表 **ColorFormat** 对象的变量。

**说明**

可以为 **Brightness** 属性输入 -1（最暗）到 1（最亮）之间的数字，0 为中间值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 设置当前文档第一个对象的前景色以及背景色和亮度 */ function test() { 	let shps = Application.ActiveDocument.Shapes; 	if (shps.Count >= 1) { 		let shp = shps.Item(1); 		let fillFmt = shp.Fill; 		fillFmt.ForeColor.RGB = 0x00ff00; 		fillFmt.BackColor.RGB = 0x0000ff; 		fillFmt.BackColor.Brightness = 0; 		fillFmt.ForeColor.Brightness = 0.4; 	} }` |



#### **ColorFormat.ObjectThemeColor**

返回或设置一个 **WdThemeColorIndex** 常量，该常量表示颜色格式的主题颜色。可读写。

**语法**

**express.ObjectThemeColor**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.RGB**

返回或设置指定颜色的 RGB 值（红-绿-蓝值），**Number** 类型，可读写。

**语法**

**express.RGB**

*express*   一个代表 **ColorFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将活动文档的第一个图形的颜色设置为灰色 */ Application.ActiveDocument.Shapes.Item(1).Fill.ForeColor.RGB = 0x7f7f7f` |

#### **ColorFormat.Type**

返回或设置形状的颜色类型。**MsoColorType** 类型，只读。

**语法**

**express.Type**

*express*   一个代表 **ColorFormat** 对象的变量。

适用环境：web

适用平台：windows/linux