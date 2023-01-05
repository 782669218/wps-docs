**ColorFormat**



代表单色对象的颜色、带有渐变或图案填充的对象的前景或背景色，或者指针的颜色。

**说明**

可以将颜色设为显式的红-绿-蓝值（使用 **RGB** 属性），或设为配色方案中的一种颜色（使用 **SchemeColor** 属性）。

使用下表中列出的属性之一可返回 **ColorFormat** 对象。

| 使用此属性         | 对象             | 返回一个 ColorFormat 对象，该对象代表          |
| ------------------ | ---------------- | ---------------------------------------------- |
| **BackColor**      | **FillFormat**   | 背景填充色（用于阴影或图案填充格式）           |
| **ForeColor**      | **FillFormat**   | 前景填充色（对于纯色填充格式，即代表填充颜色） |
| **BackColor**      | **LineFormat**   | 线条背景色（用于图案线条）                     |
| **ForeColor**      | **LineFormat**   | 线条前景色（对于纯色线条，即代表线条颜色）     |
| **ForeColor**      | **ShadowFormat** | 阴影颜色                                       |
| **ExtrusionColor** | **ThreeDFormat** | 有延伸的对象的侧边颜色                         |

使用 **RGB** 属性可将颜色设置为显示的红-绿-蓝值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向 myDocument 中添加一个矩形，然后设置矩形填充的前景色、背景色和渐变。*/ function test(){ let myDocument = Worksheets.Item(1) let myFill=  myDocument.Shapes.AddShape(msoShapeRectangle,          90, 90, 90, 50).Fill     myFill.ForeColor.RGB = (128, 0, 0)     myFill.BackColor.RGB = (170, 170, 170)     myFill.TwoColorGradient(msoGradientHorizontal, 1) }` |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Brightness**       | 返回或设置指定对象的发光度。可读/写。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ObjectThemeColor** | 返回或设置一种映射到主题配色方案的颜色。可读/写 **MsoThemeColorIndex**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RGB**              | 返回或设置一个 **Long** 值，它代表指定颜色的红绿蓝 (RGB) 值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SchemeColor**      | 返回或设置一个 **Integer** 值，它代表 **Color** 对象的颜色，作为当前配色方案中的索引。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade**     | 返回或设置一个 **Single**，使颜色变深或变浅。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**             | 返回一个代表颜色格式类型的 **MsoColorType** 值。             |

**成员属性**

#### **ColorFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ColorFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ let myObject = ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") } }` |

#### **ColorFormat.Brightness**

返回或设置指定对象的发光度。可读/写。

**语法**

**express.Brightness**

*express*   一个代表 **ColorFormat** 对象的变量。

**说明**

该属性的值必须是 0.0（最暗）到 1.0（最亮）之间的数字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例设置活动工作表中第一个形状的填充色的亮度。*/ Application.ActiveSheet.Shapes.Item(1).Fill.ForeColor.Brightness = 0.5 ` |

#### **ColorFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.ObjectThemeColor**

返回或设置一种映射到主题配色方案的颜色。可读/写 **MsoThemeColorIndex**。

**语法**

**express.ObjectThemeColor**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.RGB**

返回或设置一个 **Long** 值，它代表指定颜色的红绿蓝 (RGB) 值。

**语法**

**express.RGB**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.SchemeColor**

返回或设置一个 **Integer** 值，它代表 **Color** 对象的颜色，作为当前配色方案中的索引。

**语法**

**express.SchemeColor**

*express*   一个代表 **ColorFormat** 对象的变量。

#### **ColorFormat.TintAndShade**

返回或设置一个 **Single**，使颜色变深或变浅。

**语法**

**express.TintAndShade**

*express*   一个代表 **ColorFormat** 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

#### **ColorFormat.Type**

返回一个代表颜色格式类型的 **MsoColorType** 值。

**语法**

**express.Type**

*express*   一个代表 **ColorFormat** 对象的变量。

适用环境：web

适用平台：windows/linux