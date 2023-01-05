#### **InlineShape**



代表文档的文字层中的对象。内嵌形状只能是图片、OLE 对象或 ActiveX 控件。**InlineShape** 对象是 **InlineShapes** 集合的一个成员。**InlineShapes** 集合包含文档、范围或所选内容中的所有内嵌形状。

**说明**

**InlineShape** 对象被视为字符，并作为字符置于文本行中。

使用 **InlineShapes**(*Index*) 可返回一个 **InlineShape** 对象，其中 *Index* 是索引号。内嵌形状没有名称。以下示例激活活动文档中的第一个内嵌形状。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.InlineShapes.Item(1).Activate()` |

Shape

 对象锁定到某一文本范围，但可以自由浮动，并且可以放置在页面上的任何位置。可以使用 

ConvertToInlineShape

 方法和 

ConvertToShape

 方法来转换形状的类型。只能将图片、OLE 对象和 ActiveX 控件转换为内嵌形状。使用 

Type

 属性可返回内嵌形状的类型：图片、链接图片、嵌入的 OLE 对象、链接的 OLE 对象或者 ActiveX 控件。

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToShape** | 将嵌入式图形转换为可自由浮动的图形。返回一个 **Shape** 对象，该对象代表新图形。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**         | 删除指定的嵌入式图形。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**          | 删除对内嵌形状所做的更改。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**         | 选择指定的内嵌形状。                                         |

**属性**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlternativeText**      | 返回或设置一个 **String** 类型的值，该值代表与网页中图形相关联的可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**          | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**              | 返回一个 **Borders** 集合，该集合代表指定图形的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Chart**                | 返回一个 **Chart** 对象，该对象代表文档中内嵌形状集合内的图表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**              | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Field**                | 返回一个 **Field** 对象，该对象代表与指定内嵌形状相关联的域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                 | 返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                 | 返回一个 **GlowFormat** 对象，该对象代表发光效果的格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GroupItems**           | 返回一个 **GroupShapes** 集合，该集合代表针对内嵌形状分组到一起的形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChart**             | 如果指定的形状是图表，则为 **True**。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasSmartArt**          | 如果形状上显示 SmartArt 图表则返回 **True**。只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**               | 返回或设置内嵌形状的高度。**Single** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalLineFormat** | 返回一个 **HorizontalLineFormat** 对象，该对象包含指定的 **InlineShape** 对象的水平行的格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlink**            | 返回一个 **Hyperlink** 对象，该对象代表与指定内嵌形状相关联的超链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsPictureBullet**      | 如果为 **True**，则表明 **InlineShape** 对象是图片项目符号。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**                 | 返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkFormat**           | 返回一个 **LinkFormat** 对象，该对象代表链接到文件的指定内嵌形状的链接选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockAspectRatio**      | 如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEFormat**            | 返回一个 **OLEFormat** 对象，该对象代表指定的内嵌形状的 OLE 特征（链接除外）。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**               | 返回一个 **Object** 类型值，该值代表指定 **InlineShape** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureFormat**        | 返回一个 **PictureFormat** 对象，该对象包含内嵌形状的图片格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                | 返回一个 **Range** 对象，该对象代表内嵌形状中包含的文档部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reflection**           | 返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScaleHeight**          | 参照指定内嵌形状的原始尺寸缩放其高度。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScaleWidth**           | 参照指定内嵌形状的原始尺寸缩放其宽度。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Script**               | 返回一个 **Script** 对象，该对象代表与指定网页中的图像相关联的脚本或代码块。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**               | 返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartArt**             | 返回 SmartArt 对象，该对象提供一种处理与指定内嵌形状相关联的 SmartArt 的方式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SoftEdge**             | 返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEffect**           | 返回一个 **TextEffectFormat** 对象，该对象包含指定内嵌形状的文本效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**                | 返回或设置包含指定内嵌形状的标题的 **String** 类型值。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                 | 返回内嵌形状的类型。只读 [WdInlineShapeType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdInlineShapeType%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                | 返回或设置指定内嵌形状的宽度（以磅为单位）。可读写 **Long** 类型。 |

**成员方法**

#### **InlineShape.ConvertToShape**

将嵌入式图形转换为可自由浮动的图形。返回一个 **Shape** 对象，该对象代表新图形。

**语法**

**express.ConvertToShape()**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

使用 **ConvertToShape** 方法前，最少应该已经对 **FreeformBuilder** 对象用过一次 **AddNodes** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的第一个嵌入式图形转化为浮动的图形。*/ ActiveDocument.InlineShapes.Item(1).ConvertToShape()` |

#### **InlineShape.Delete**

删除指定的嵌入式图形。

**语法**

**express.Delete()**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.Reset**

删除对内嵌形状所做的更改。

**语法**

**express.Reset()**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先将一幅图片作为内嵌形状插入，并更改其亮度，然后再将其亮度重新设置为初始的亮度。*/ function test() { let aInLine = ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Bubbles.bmp", Selection.Range) aInLine.PictureFormat.Brightness = 0.5 MsgBox("Changing brightness back") aInLine.Reset() }` |

#### **InlineShape.Select**

选择指定的内嵌形状。

**语法**

**express.Select()**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

使用此方法之后，可用 **Selection** 属性处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

**成员属性**

#### **InlineShape.AlternativeText**

返回或设置一个 **String** 类型的值，该值代表与网页中图形相关联的可选文字?（可选文字：由 Web 浏览器用于在图像下载过程中显示文本，其对象是那些关闭图形的用户，以及那些依赖屏幕阅读软件将屏幕上的图形转换为可读语言的用户。）。可读写。

**语法**

**express.AlternativeText**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：为活动窗口中选定的图形设定可选文字。选定图形是一幅野鸭的图片。*/ ActiveWindow.Selection.InlineShapes.Item(1).AlternativeText = "This is a mallard duck."` |

#### **InlineShape.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **InlineShape.Borders**

返回一个 **Borders** 集合，该集合代表指定图形的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **InlineShape.Chart**

返回一个 **Chart** 对象，该对象代表文档中内嵌形状集合内的图表。只读。

**语法**

**express.Chart**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **InlineShape.Field**

返回一个 **Field** 对象，该对象代表与指定内嵌形状相关联的域。只读。

**语法**

**express.Field**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

使用 **Fields** 属性可返回 **Fields** 集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将一个图形作为内嵌形状插入（使用 INCLUDEPICTURE 域），然后显示该形状的域代码。*/ function test() { let iShapeNew =ActiveDocument.InlineShapes.AddPicture("C:\\Windows\\Tiles.bmp",true,false,Selection.Range) MsgBox(ActiveDocument.InlineShapes.Item(1).Field.Code.Text) }` |

#### **InlineShape.Fill**

返回一个 **FillFormat** 对象，该对象包含指定图形的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先向 myDocument 添加一个矩形，然后为其设置前景色、背景色和矩形填充的渐变类型。*/ function test() { let myDocument = Documents.Item(1).Shapes.AddShape(msoShapeRectangle,90, 90, 90, 50).Fill     myDocument.ForeColor.RGB = (128, 0, 0)     myDocument.BackColor.RGB = (170, 170, 170)     myDocument.TwoColorGradient(msoGradientHorizontal, 1) }` |

#### **InlineShape.Glow**

返回一个 **GlowFormat** 对象，该对象代表发光效果的格式属性。只读。

**语法**

**express.Glow**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.GroupItems**

返回一个 **GroupShapes** 集合，该集合代表针对内嵌形状分组到一起的形状。只读。

**语法**

**express.GroupItems**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.HasChart**

如果指定的形状是图表，则为 **True**。只读。

**语法**

**express.HasChart**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

对于 OLE 图表，此属性总是返回 False。此时请使用 `InlineShape.OLEFormat.ProgID`，并且检查是否存在下列可能的值：“Excel.Chart.8”、“MSGraph.Chart.8”、“Excel.Sheet.8”、“Excel.Chart.5”、“MSGraph.Chart.5”或“Excel.Sheet.5”。

#### **InlineShape.HasSmartArt**

如果形状上显示 SmartArt 图表则返回 **True**。只读。

**语法**

**express.HasSmartArt**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中的第一个内嵌形状是否包含 SmartArt。*/ function test() { let myInlineShape = ActiveDocument.InlineShapes.Item(1)  if(myInlineShape.HasSmartArt){    MsgBox("The first shape contains SmartArt.") } else{    MsgBox("The first shape contains no SmartArt.") } }` |

#### **InlineShape.Height**

返回或设置内嵌形状的高度。**Single** 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.HorizontalLineFormat**

返回一个 **HorizontalLineFormat** 对象，该对象包含指定的 **InlineShape** 对象的水平行的格式。只读。

**语法**

**express.HorizontalLineFormat**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将指定水平行的长度设定为窗口宽度的 50%。*/ function test() { ActiveDocument.InlineShapes.AddHorizontalLineStandard() ActiveDocument.InlineShapes.Item(1).HorizontalLineFormat.PercentWidth = 50 }` |

#### **InlineShape.Hyperlink**

返回一个 **Hyperlink** 对象，该对象代表与指定内嵌形状相关联的超链接。只读。

**语法**

**express.Hyperlink**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

如果不存在与指定形状相关联的超链接，则会发生错误。在此情况下，请将 **Add** 方法用于 **Hyperlinks** 集合来为指定形状添加超链接。以下示例说明了具体的操作方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `ActiveDocument.Hyperlinks.Add(Selection.InlineShapes.Item(1), "http://www.wps.cn")    ` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中第一个形状的超链接地址。*/ MsgBox(ActiveDocument.Shapes.Item(1).Hyperlink.Address)` |

#### **InlineShape.IsPictureBullet**

如果为 **True**，则表明 **InlineShape** 对象是图片项目符号。**Boolean** 类型，只读。

**语法**

**express.IsPictureBullet**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

虽然图片项目符号被看作嵌入式图形，但是搜索文档的 **InlineShapes** 集合并不返回图片项目符号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果选定列表的格式被设为图片项目符号，本示例设置该列表的格式，否则，显示一条消息。*/ function IsSelectionAPictureBullet(shp){     try{         if(shp.IsPictureBullet == true){           shp.Width = InchesToPoints(0.5)           shp.Height = InchesToPoints(0.05)         }     }     catch(exception){           MsgBox("The selection is not a list or " + "does not contain picture bullets.")     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用下列代码调用以上程序。*/ function CallPic(){     IsSelectionAPictureBullet(Selection.Range.ListFormat.ListPictureBullet) }` |

#### **InlineShape.Line**

返回一个 **LineFormat** 对象，该对象包含指定形状的线条格式属性。只读。

**语法**

**express.Line**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

对于有边框的内嵌形状，**LineFormat** 对象代表边框。

#### **InlineShape.LinkFormat**

返回一个 **LinkFormat** 对象，该对象代表链接到文件的指定内嵌形状的链接选项。只读。

**语法**

**express.LinkFormat**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.LockAspectRatio**

如果在调整指定形状的大小时保留其最初比例，则该属性值为 **MsoTrue**；如果在调整形状大小时可分别改变其高度和宽度，则该属性值为 **MsoFalse**。可读写 **MsoTriState** 类型。

**语法**

**express.LockAspectRatio**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.OLEFormat**

返回一个 **OLEFormat** 对象，该对象代表指定的内嵌形状的 OLE 特征（链接除外）。只读

**语法**

**express.OLEFormat**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.Parent**

返回一个 **Object** 类型值，该值代表指定 **InlineShape** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.PictureFormat**

返回一个 **PictureFormat** 对象，该对象包含内嵌形状的图片格式属性。只读。

**语法**

**express.PictureFormat**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

适用于代表图片或 OLE 对象的 **InlineShape** 对象。

#### **InlineShape.Range**

返回一个 **Range** 对象，该对象代表内嵌形状中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.Reflection**

返回一个 **ReflectionFormat** 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.ScaleHeight**

参照指定内嵌形状的原始尺寸缩放其高度。可读写 **Single** 类型。

**语法**

**express.ScaleHeight**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个内嵌形状的高度和宽度设为其原高度和宽度的 150%。*/ function test() { ActiveDocument.InlineShapes.Item(1).ScaleHeight = 150 ActiveDocument.InlineShapes.Item(1).ScaleWidth = 150 }  ` |

#### **InlineShape.ScaleWidth**

参照指定内嵌形状的原始尺寸缩放其宽度。可读写 **Single** 类型。

**语法**

**express.ScaleWidth**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个内嵌形状的高度和宽度设为其原高度和宽度的 150%。*/ function test() { ActiveDocument.InlineShapes.Item(1).ScaleHeight = 150 ActiveDocument.InlineShapes.Item(1).ScaleWidth = 150 }` |

#### **InlineShape.Script**

返回一个 **Script** 对象，该对象代表与指定网页中的图像相关联的脚本或代码块。

**语法**

**express.Script**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

如果该网页中不包含任何脚本，则不返回任何内容。

#### **InlineShape.Shadow**

返回一个 **ShadowFormat** 对象，该对象代表指定形状的阴影格式。只读。

**语法**

**express.Shadow**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.SmartArt**

返回 SmartArt 对象，该对象提供一种处理与指定内嵌形状相关联的 SmartArt 的方式。只读。

**语法**

**express.SmartArt**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

**SmartArt** 属性为与内嵌形状相关联的 SmartArt 图形进行交互提供了一个入口点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中添加 SmartArt 图形。*/ function test() { let myDoc = ActiveDocument let myInlineShape = myDoc.InlineShapes.AddSmartArt(Application.SmartArtLayouts.Item(2), null, null, null, null, myDoc.Paragraphs.Item(2).Range) let mySmartArt = myInlineShape.SmartArt }` |

#### **InlineShape.SoftEdge**

返回一个 **SoftEdgeFormat** 对象，该对象代表形状的软边缘格式。只读。

**语法**

**express.SoftEdge**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.TextEffect**

返回一个 **TextEffectFormat** 对象，该对象包含指定内嵌形状的文本效果格式属性。只读。

**语法**

**express.TextEffect**

*express*   一个代表 **InlineShape** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，如果 myDocument 的第三个形状为艺术字，则将它的字形设为加粗。*/ function test() { let myDocument = ActiveDocument let myShapes = myDocument.Shapes.Item(3)  if(myShapes.Type == msoTextEffect){     myShapes.TextEffect.FontBold = true } }` |

#### **InlineShape.Title**

返回或设置包含指定内嵌形状的标题的 **String** 类型值。可读/写。

**语法**

**express.Title**

*express*   一个代表 **InlineShape** 对象的变量。

**说明**

使用 **Title** 属性可为内嵌形状提供可选文字标题。该属性将标题文本添加到 WPS 2015 中**“设置形状格式”**对话框的**“可选文字”**窗格上的**“标题”**文本框中。

| 注释                                                         |
| ------------------------------------------------------------ |
| Web 浏览器在加载表格的过程中或表格丢失时显示可选文字。Web 搜索引擎利用可选文字帮助查找网页。可选文字也可用来帮助残障人士。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档的第一个内嵌形状中添加可选文字标题。 */ ActiveDocument.InlineShapes.Item(1).Title = "Desktop screenshot."` |

#### **InlineShape.Type**

返回内嵌形状的类型。只读 [WdInlineShapeType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdInlineShapeType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **InlineShape** 对象的变量。

#### **InlineShape.Width**

返回或设置指定内嵌形状的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **InlineShape** 对象的变量。

适用环境：web

适用平台：windows/linux