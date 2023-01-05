**LineFormat**



代表线条和箭头格式。

**说明**

对于线条，**LineFormat** 对象包含该线条自身的格式信息；对于有边界的形状，该对象包含形状的边界的格式信息。

使用 **Line** 属性可返回一个 **LineFormat** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给 myDocument 添加一条蓝色虚线。在该线的起点有一个短而窄的椭圆，在该线的终点有一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.DashStyle = msoLineDashDotDot line.ForeColor.RGB = (50, 0, 128) line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide } ` |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackColor**            | 返回或设置一个 **ColorFormat** 对象，它代表指定的填充背景色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadLength** | 返回或设置指定线条起点的箭头的长度。**MsoArrowheadLength** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadStyle**  | 返回或设置指定线条起点的箭头样式。**MsoArrowheadStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginArrowheadWidth**  | 返回或设置指定线条起点的箭头宽度。**MsoArrowheadWidth** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**              | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DashStyle**            | 返回或设置指定直线的虚线样式。可以是 **MsoLineDashStyle** 常量之一。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadLength**   | 返回或设置指定线条末尾的箭头的长度。**MsoArrowheadLength** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadStyle**    | 返回或设置指定线条端点的箭头样式。**MsoArrowheadStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndArrowheadWidth**    | 返回或设置指定线条端点的箭头的宽度。**MsoArrowheadWidth** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ForeColor**            | 返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsetPen**             | 返回或设置线条是否绘制在指定形状的边界以内。可读/写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pattern**              | 返回或设置一个代表填充图案的 **MsoPatternType** 值。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                | 返回或设置一个 **MsoLineStyle** 值，它代表线条的样式。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Transparency**         | 返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。**Double** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**              | 返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Weight**               | 返回或设置一个 **Single** 值，它代表线条的粗细。             |

**成员属性**

#### **LineFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **LineFormat.BackColor**

返回或设置一个 **ColorFormat** 对象，它代表指定的填充背景色。

**语法**

**express.BackColor**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.BeginArrowheadLength**

返回或设置指定线条起点的箭头的长度。**MsoArrowheadLength** 类型，可读写。

**语法**

**express.BeginArrowheadLength**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.BeginArrowheadStyle**

返回或设置指定线条起点的箭头样式。**MsoArrowheadStyle** 类型，可读写。

**语法**

**express.BeginArrowheadStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.BeginArrowheadWidth**

返回或设置指定线条起点的箭头宽度。**MsoArrowheadWidth** 类型，可读写。

**语法**

**express.BeginArrowheadWidth**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LineFormat.DashStyle**

返回或设置指定直线的虚线样式。可以是 **MsoLineDashStyle** 常量之一。可读/写 **Long** 类型。

**语法**

**express.DashStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加蓝色的虚线。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(10, 10, 250, 250).Line line.DashStyle = msoLineDashDotDot line.ForeColor.RGB = RGB(50, 0, 128) }` |

#### **LineFormat.EndArrowheadLength**

返回或设置指定线条末尾的箭头的长度。**MsoArrowheadLength** 类型，可读写。

**语法**

**express.EndArrowheadLength**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| **MsoArrowheadLength** 可以是下列 **MsoArrowheadLength** 常量之一。 |
| ------------------------------------------------------------ |
| **msoArrowheadLengthMixed**                                  |
| **msoArrowheadShort**                                        |
| **msoArrowheadLengthMedium**                                 |
| **msoArrowheadLong**                                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWideh }` |

#### **LineFormat.EndArrowheadStyle**

返回或设置指定线条端点的箭头样式。**MsoArrowheadStyle** 类型，可读写。

**语法**

**express.EndArrowheadStyle**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| **MsoArrowheadStyle** 可以是下列 **MsoArrowheadStyle** 常量之一。 |
| ------------------------------------------------------------ |
| **msoArrowheadNone**                                         |
| **msoArrowheadOval**                                         |
| **msoArrowheadStyleMixed**                                   |
| **msoArrowheadDiamond**                                      |
| **msoArrowheadOpen**                                         |
| **msoArrowheadStealth**                                      |
| **msoArrowheadTriangle**                                     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Application.Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.EndArrowheadWidth**

返回或设置指定线条端点的箭头的宽度。**MsoArrowheadWidth** 类型，可读写。

**语法**

**express.EndArrowheadWidth**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

| **MsoArrowheadWidth** 可以是下列 **MsoArrowheadWidth** 常量之一。 |
| ------------------------------------------------------------ |
| **msoArrowheadNarrow**                                       |
| **msoArrowheadWidthMedium**                                  |
| **msoArrowheadWide**                                         |
| **msoArrowheadWidthMixed**                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加直线。该直线的起点是一个短而窄的椭圆，而该直线的终点则是一个长而宽的三角形。*/ function test(){ let myDocument = Worksheets.Item(1) let line = myDocument.Shapes.AddLine(100, 100, 200, 300).Line line.BeginArrowheadLength = msoArrowheadShort line.BeginArrowheadStyle = msoArrowheadOval line.BeginArrowheadWidth = msoArrowheadNarrow line.EndArrowheadLength = msoArrowheadLong line.EndArrowheadStyle = msoArrowheadTriangle line.EndArrowheadWidth = msoArrowheadWide }` |

#### **LineFormat.ForeColor**

返回或设置一个 **ColorFormat** 对象，它代表指定的前景填充色或纯色。

**语法**

**express.ForeColor**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.InsetPen**

返回或设置线条是否绘制在指定形状的边界以内。可读/写。

**语法**

**express.InsetPen**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

如果线条绘制在形状的边界以内，则为 **msoTrue** (-1)；否则为 **msoFalse** (0)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将两个矩形添加到活动工作表，第一个矩形的线条绘制在其边界以内，第二个矩形的线条绘制在其边界上。*/ function test(){ let shpNew = Application.ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 150, 150, 100) shpNew.Line.Weight = 24 shpNew.Line.InsetPen = msoTrue  let shpNew2 = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 200, 150, 150, 100) shpNew2.Line.Weight = 24 shpNew2.Line.InsetPen = msoFalse }` |

#### **LineFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.Pattern**

返回或设置一个代表填充图案的 **MsoPatternType** 值。

**语法**

**express.Pattern**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.Style**

返回或设置一个 **MsoLineStyle** 值，它代表线条的样式。

**语法**

**express.Style**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.Transparency**

返回或设置指定填充的透明度，取值范围为 0.0（不透明）到 1.0（清晰）之间。**Double** 型，可读写。

**语法**

**express.Transparency**

*express*   一个代表 **LineFormat** 对象的变量。

**说明**

此属性的值仅影响均一颜色填充和线条的外观。对图案线条及渐变、图案、图片和纹理填充的外观无效。

#### **LineFormat.Visible**

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **LineFormat** 对象的变量。

#### **LineFormat.Weight**

返回或设置一个 **Single** 值，它代表线条的粗细。

**语法**

**express.Weight**

*express*   一个代表 **LineFormat** 对象的变量。

适用环境：web

适用平台：windows/linux