#### **InlineShapes**



**InlineShape** 对象的集合，这些对象代表文档、范围或所选内容中的所有内嵌形状。

**方法**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddChart**                  | 将指定类型的图表作为内嵌形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddHorizontalLine**         | 向当前文档添加一条基于图像文件的横线。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddHorizontalLineStandard** | 向当前文档添加一条横线。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEControl**             | 创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEObject**              | 创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPicture**                | 在文档中添加一幅图片。返回一个代表该图片的 [InlineShape ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/InlineShape/InlineShape%20.htm#jsObject_InlineShape)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPictureBullet**          | 向当前文档添加一个基于图像文件的图片项目符号。返回一个 **InlineShape** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddSmartArt**               | 将 SmartArt 图形作为内嵌形状插入活动文档。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                      | 返回集合中的单个 **InlineShape** 对象。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **New**                       | 插入一个空白的 WPS 图片对象，该对象是边长为一英寸的正方形，四周带有边框。此方法以 **InlineShape** 对象的形式返回新图形。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中内嵌图形的数目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **InlineShapes** 对象的父对象。 |

**成员方法**

#### **InlineShapes.AddChart**

将指定类型的图表作为内嵌形状插入到活动文档中，并且在ET 中打开包含 WPS 用于创建该图表的默认数据的工作表。

**语法**

**express.AddChart(Type, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                                     |
| -------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Type*   | 必选          | **XlChartType** | 指定要创建的图表的类型。                                     |
| *Range*  | 必选          | **Variant**     | 指定图表所绑定到的文本。如果指定了 Range，则会将图表放置在该区域中第一段的开头。如果省略该参数，则自动选择该区域，并且相对于页面的上边缘和左边缘放置图表。 |

**返回值**

InlineShape

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*创建一个新的三维柱形图，然后将其添加到活动文档中。*/ ActiveDocument.InlineShapes.AddChart(xl3DColumn)` |

#### **InlineShapes.AddHorizontalLine**

向当前文档添加一条基于图像文件的横线。

**语法**

**express.AddHorizontalLine(FileName, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName* | 必选          | **String**   | 要用于横线的图像的文件名。                                   |
| *Range*    | 可选          | **Variant**  | WPS 在其上方放置横线的区域。如果省略该参数，则 WPS 会将横线放置在当前选定内容的上方。 |

**说明**

要添加不基于现有图像文件的横线，可用 **AddHorizontalLineStandard** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一条基于图像文件“ArtsyRule.gif”的横线添加到活动文档中当前选定内容的上方。*/ Selection.InlineShapes.AddHorizontalLine("C:\\Art files\\ArtsyRule.gif")    ` |

#### **InlineShapes.AddHorizontalLineStandard**

向当前文档添加一条横线。

**语法**

**express.AddHorizontalLineStandard(Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Range*  | 可选          | **Variant**  | WPS 在其上方放置横线的区域。如果省略该参数，则 WPS 会将横线放置在当前选定内容的上方。 |

**说明**

要添加基于现有图像文件的横线，可用 **AddHorizontalLine** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一条横线添加到活动文档第五段的上方。*/ ActiveDocument.Paragraphs.Item(5).Range.InlineShapes.AddHorizontalLineStandard()    ` |

#### **InlineShapes.AddOLEControl**

创建一个 ActiveX 控件（以前称为 OLE 控件）。返回代表新 ActiveX 控件的 **InlineShape** 对象。

**语法**

**express.AddOLEControl(ClassType, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType* | 可选          | **Variant**  | 用于要创建的 ActiveX 控件的编程标识符。                      |
| *Range*     | 可选          | **Variant**  | 一个区域，ActiveX 控件将放置在该区域的文本中。如果该区域未折叠，则 ActiveX 控件将替换该区域。如果省略该参数，则自动放置 ActiveX 控件。 |

**说明**

在 WPS 中，将 ActiveX 控件表示为 **Shape** 对象或 **InlineShape** 对象。若要修改 ActiveX 控件的属性，可以对指定的图形或嵌入图形使用 **OLEFormat** 对象的 **Object** 属性。

关于 ActiveX 控件类型的可用信息，请参阅 OLE 编程标识符

#### **InlineShapes.AddOLEObject**

创建一个 OLE 对象。返回代表新 OLE 对象的 **InlineShape** 对象。

**语法**

**express.AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | 用于激活指定 OLE 对象的应用程序名称。                        |
| *FileName*      | 可选          | **Variant**  | 创建对象所用的文件。如果省略该参数，则使用当前的文件夹。必须指定对象的 ClassType 或 FileName 参数，但不能同时指定两者。 |
| *LinkToFile*    | 可选          | **Variant**  | 如果为 True，则将 OLE 对象链接到创建该对象所用的文件。如果为 False，则 OLE 对象成为其源文件的独立副本。如果为 ClassType 指定了值，则 LinkToFile 参数必须为 False。默认值为 False。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果为 True，则将 OLE 对象显示为图标。默认值为 False。       |
| *IconFileName*  | 可选          | **Variant**  | 要显示的图标所在的文件。                                     |
| *IconIndex*     | 可选          | **Variant**  | IconFileName 中图标的索引号。当选中“显示为图标”复选框时，在指定文件中图标的顺序对应于“插入”菜单“对象”对话框的“更改图标”对话框中图标出现的顺序。文件中的第一个图标的索引号为 0（零）。如果指定索引号的图标在 IconFileName 中不存在，则使用索引号为 1 的图标（文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选          | **Variant**  | 显示在图标下面的标签（题注）。                               |
| *Range*         | 可选          | **Variant**  | 指定 OLE 对象放置在文本中的区域。除非区域是折叠的，否则 OLE 对象将替换该区域。如果省略该参数，则自动放置对象。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的第二段中添加一张新的 ET 工作表。*/ function test() { ActiveDocument.InlineShapes.AddOLEObject("Excel.Sheet", null, null, false,     null, null, null, ActiveDocument.Paragraphs.Item(2).Range) }` |

#### **InlineShapes.AddPicture**

在文档中添加一幅图片。返回一个代表该图片的 [InlineShape ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/InlineShape/InlineShape%20.htm#jsObject_InlineShape)对象。

**语法**

**express.AddPicture(FileName, LinkToFile, SaveWithDocument, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*         | 必选          | **String**   | 图片的路径和文件名。                                         |
| *LinkToFile*       | 可选          | **Variant**  | 如果为 True，则将图片链接到创建它的文件；如果为 False，则使图片成为该文件的独立副本。默认值为 False。 |
| *SaveWithDocument* | 可选          | **Variant**  | 如果为 True，则将链接的图片与文档一起保存。默认值为 False。  |
| *Range*            | 可选          | **Variant**  | 图片置于文本中的位置。如果该区域未折叠，那么图片将覆盖此区域，否则插入图片。如果省略此参数，则自动放置图片。 |

#### **InlineShapes.AddPictureBullet**

向当前文档添加一个基于图像文件的图片项目符号。返回一个 **InlineShape** 对象。

**语法**

**express.AddPictureBullet(FileName, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName* | 必选          | **String**   | 要用于图片项目符号的图像的文件名称。                         |
| *Range*    | 可选          | **Variant**  | WPS 向其添加图片项目符号的区域。 WPS 会将图片项目符号添加到该区域的每个段落。如果省略该参数，则 WPS 将图片项目符号添加到当前选定内容的每个段落中。 |

**返回值**

InlineShape

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使用名为“RedBullet.gif”的图像文件向选定文本的每个段落添加一个图片项目符号。*/ Selection.InlineShapes.AddPictureBullet("C:\\Art files\\RedBullet.gif")    ` |

#### **InlineShapes.AddSmartArt**

将 SmartArt 图形作为内嵌形状插入活动文档。

**语法**

**express.AddSmartArt(Layout, Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**         | **说明**                                                     |
| -------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Layout* | 必选          | **[SMARTARTLAYOUT]** | 指定 SmartArt 图形的布局的 SmartArtLayout 对象。             |
| *Range*  | 可选          | **Variant**          | 指定 SmartArt 图形所绑定到的文本。如果指定了 Range，则会将 SmartArt 图形置于该范围中第一段的开头。如果省略该参数，则自动选择该范围，并且相对于页面的上边缘和左边缘放置 SmartArt 图形。 |

**返回值**

InlineShape

#### **InlineShapes.Item**

返回集合中的单个 **InlineShape** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

#### **InlineShapes.New**

插入一个空白的 WPS 图片对象，该对象是边长为一英寸的正方形，四周带有边框。此方法以 **InlineShape** 对象的形式返回新图形。

**语法**

**express.New(Range)**

*express*   一个代表 **InlineShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**       |
| -------- | ------------- | ---------------- | -------------- |
| *Range*  | 必选          | **Range object** | 新图形的位置。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中插入一个新的空白图片，并在该图片周围应用阴影边框。*/ function test() { let ishapeNew = ActiveDocument.InlineShapes.New(Selection.Range)     ishapeNew.Borders.Shadow = true     ActiveDocument.ActiveWindow.View.ShowFieldCodes = false }` |

**成员属性**

#### **InlineShapes.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **InlineShapes** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **InlineShapes.Count**

返回一个 **Long** 类型的值，该值代表集合中内嵌图形的数目。只读。

**语法**

**express.Count**

*express*   一个代表 **InlineShapes** 对象的变量。

#### **InlineShapes.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **InlineShapes** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **InlineShapes.Parent**

返回一个 **Object** 类型值，该值代表指定 **InlineShapes** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **InlineShapes** 对象的变量。

适用环境：web

适用平台：windows/linux