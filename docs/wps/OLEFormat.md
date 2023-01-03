#### **OLEFormat**



代表 OLE 对象、ActiveX 控件或域的 OLE（而不是链接）特性。

**说明**

使用形状、内嵌形状或域的 **OLEFormat** 属性可返回 **OLEFormat** 对象。以下示例显示活动文档中的第一个形状的类类型。

| 示例代码复制                                                |
| ----------------------------------------------------------- |
| `MsgBox(ActiveDocument.Shapes.Item(1).OLEFormat.ClassType)` |

并非所有类型的形状、内嵌形状和域都具有 OLE 功能。使用 **Shape** 和 **InlineShape** 对象的 **Type** 属性可确定指定形状或内嵌形状的类别。**Field** 对象的 **Type** 属性返回域的类型。

您可以使用 **Activate**、**Edit**、**Open** 和 **DoVerb** 方法自动处理 OLE 对象。

使用 **Object** 属性可返回一个代表 ActiveX 控件或 OLE 对象的对象。借助于该对象，您可以使用容器应用程序或 ActiveX 控件的属性和方法。

**方法**

|                                                              | 名称           | 说明                                                         |
| ------------------------------------------------------------ | -------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**   | 激活指定的 **OLEFormat** 对象。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ActivateAs** | 设置确定用于激活指定 OLE 对象的默认应用程序的 Windows 注册表值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertTo**  | 将指定的 OLE 对象转换为另一种类型，以便能在另一种服务器应用程序中编辑该对象，或者改变对象在文档中的显示方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DoVerb**     | 请求 OLE 对象执行一个有效动作，以激活其内容。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Edit**       | 在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Open**       | 打开指定的 **OLEFormat** 对象。                              |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ClassType**                  | 返回或设置指定的 OLE 对象、图片或域的类型。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAsIcon**              | 如果指定对象显示为图标，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IconIndex**                  | 返回或设置 **DisplayAsIcon** 属性为 **True** 时所用的图标。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IconLabel**                  | 返回或设置 OLE 对象的图标下面显示的文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IconName**                   | 返回或设置存储 OLE 对象的图标的程序文件。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IconPath**                   | 返回存储 OLE 对象的图标的文件的路径。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Label**                      | 返回用于标识源文件中被链接的部分的字符串。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Object**                     | 返回一个 **Object** 类型的值，该值代表指定的 OLE 对象的最上层接口。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回一个 **Object** 类型值，该值代表指定 **OLEFormat** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreserveFormattingOnUpdate** | 如果为 **True**，则将 WPS 中的格式设置保存到链接的 OLE 对象中，例如链接到 ET 电子表格的表格。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProgID**                     | 返回指定 OLE 对象的程序标识符 (ProgID)?（程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或 PowerPoint.Slide），由 Windows 注册表用来唯一标识一个对象。）。只读 **String** 类型。 |

**成员方法**

#### **OLEFormat.Activate**

激活指定的 **OLEFormat** 对象。

**语法**

**express.Activate()**

*express*   一个代表 **OLEFormat** 对象的变量。

#### **OLEFormat.ActivateAs**

设置确定用于激活指定 OLE 对象的默认应用程序的 Windows 注册表值。

**语法**

**express.ActivateAs(ClassType)**

*express*   一个代表 **OLEFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType* | 必选          | **String**   | 在其中打开 OLE 对象的应用程序的名称。要查看 OLE 对象可被激活为的对象类型列表，请单击该对象，然后打开“转换”对话框。将一个对象作为内嵌形状插入，然后查看域代码，就可以找到 ClassType 字符串。对象的类类型后带有“EMBED”或“LINK”一词。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置在 ET 中打开活动文档的第一个浮动图形，然后激活该图形。为使本示例能够运行，此图形必须是可以在 ET 中打开的 OLE 对象。*/ function test() { ActiveDocument.Shapes.Item(1).OLEFormat.ActivateAs("Excel.Sheet") ActiveDocument.Shapes.Item(1).OLEFormat.Activate() }` |

#### **OLEFormat.ConvertTo**

将指定的 OLE 对象转换为另一种类型，以便能在另一种服务器应用程序中编辑该对象，或者改变对象在文档中的显示方式。

**语法**

**express.ConvertTo(ClassType, DisplayAsIcon, IconFileName, IconIndex, IconLabel)**

*express*   一个代表 **OLEFormat** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | 用于激活 OLE 对象的应用程序的名称。在“对象”对话框的“新建”选项卡上的“对象类型”框中可以看到可用应用程序列表。将一个对象作为内嵌形状插入，然后查看域代码，就可以找到 ClassType 字符串。对象的类类型后带有“EMBED”或“LINK”一词。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果该属性值为 True，则将 OLE 对象显示为图标。默认值为 False。 |
| *IconFileName*  | 可选          | **Variant**  | 包含将要显示的图标的文件。                                   |
| *IconIndex*     | 可选          | **Variant**  | IconFileName 内的图标的索引编号。指定文件中图标的顺序与在选定“显示为图标”复选框后“更改图标”对话框中（可通过“插入对象”对话框访问）图标显示的顺序对应。文件中的第一个图标的索引编号为 0（零）。如果具有给定索引编号的图标在 IconFileName 所指定的文件中不存在，则使用索引编号为 1 的图标（即文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选          | **Variant**  | 显示在图标下面的标签（题注）。                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一篇文档，然后插入一篇嵌入的包含文字的文档。再将嵌入的文档转换成 WPS 图片。*/ function test() { Documents.Add() let objEmbedded = ActiveDocument.Shapes.AddOLEObject("Word.Document")     objEmbedded.Activate()  Selection.TypeText("Test") objEmbedded.OLEFormat.OLEFormat.ConvertTo("Word.Picture") }` |

#### **OLEFormat.DoVerb**

请求 OLE 对象执行一个有效动作，以激活其内容。

**语法**

**express.DoVerb(VerbIndex)**

*express*   一个代表 **OLEFormat** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *VerbIndex* | 可选          | **Variant**  | OLE 对象应执行的动作。如果省略该参数，则发送默认动作。如果 OLE 对象不支持被请求的动作，将发生错误。可以是任意 WdOLEVerb 常量。 |

**说明**

每个 OLE 对象都支持附加于该对象的一组动作。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第一个浮动 OLE 对象的默认动作发送给服务器。*/ ActiveDocument.Shapes.Item(1).OLEFormat.DoVerb()` |

#### **OLEFormat.Edit**

在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。

**语法**

**express.Edit()**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开（为了进行编辑）活动文档中第一个嵌入式 OLE 对象（定义为图形）。*/ function test() { let shapesAll = ActiveDocument.Shapes if(shapesAll.Count >= 1){     if(shapesAll.Item(1).Type == msoEmbeddedOLEObject){         shapesAll.Item(1).OLEFormat.Edit()     } } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开（为了进行编辑）活动文档中第一个链接的 OLE 对象（定义为嵌入式图形）。*/ function test() { let colIS = ActiveDocument.InlineShapes if(colIS.Count >= 1){     if(colIS.Item(1).Type == wdInlineShapeLinkedOLEObject){         colIS.Item(1).OLEFormat.Edit()     } } }` |

#### **OLEFormat.Open**

打开指定的 **OLEFormat** 对象。

**语法**

**express.Open()**

*express*   一个代表 **OLEFormat** 对象的变量。

**成员属性**

#### **OLEFormat.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **OLEFormat.ClassType**

返回或设置指定的 OLE 对象、图片或域的类型。**String** 类型，可读写。

**语法**

**express.ClassType**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

对于除 DDE 链接之外的其他链接对象，本属性是只读的。

在**“插入”**菜单上**“对象”**对话框中**“新建”**选项卡上的**“对象类型”**框中可以看到可用应用程序的列表。以嵌入式图形形式插入一个对象，然后查看域代码，可以找到 **ClassType** 字符串。对象的类类型后面都带有单词“EMBED”或“LINK”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例遍历活动文档中所有浮动图形，并将所有链接的 ET 工作表设置为可自动更新。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){         if(ActiveDocument.Shapes.Item(i).OLEFormat.ClassType == "Excel.Sheet"){             ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = true         }     } } }  ` |

#### **OLEFormat.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **OLEFormat.DisplayAsIcon**

如果指定对象显示为图标，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayAsIcon**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一个消息框，框内包含活动文档中每个显示为图标的浮动图形名称。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).OLEFormat.DisplayAsIcon == true){        MsgBox(shapeLoop.Name + " is displayed as an icon.")     } } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 ET 工作表作为链接的 OLE 对象插入活动文档，然后将该对象的显示更改为图标方式。*/ function test() { let objNew = ActiveDocument.Shapes.AddOLEObject(null, "C:\\Program Files\\Microsoft Office\\Office\\Samples\\samples.xls", true) ActiveDocument.Shapes.Item(1).OLEFormat.DisplayAsIcon = true }` |

#### **OLEFormat.IconIndex**

返回或设置 **DisplayAsIcon** 属性为 **True** 时所用的图标。**Long** 类型，可读写。

**语法**

**express.IconIndex**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

零 (0) 对应第一个图标，1 对应第二个图标，依此类推。如果省略该参数，则使用第一个（默认）图标。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在消息框中返回显示为图标的第一个选定图形的图标索引号。*/ function test() { if(Selection.ShapeRange.Count >= 1){     let olefTemp = Selection.ShapeRange.Item(1).OLEFormat     if(olefTemp.DisplayAsIcon == true){         MsgBox(olefTemp.IconIndex)     } } }` |

#### **OLEFormat.IconLabel**

返回或设置 OLE 对象的图标下面显示的文字。**String** 类型，可读写。

**语法**

**express.IconLabel**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改选定内容中第一个图形的图标下面的文字。*/ function test() { if(Selection.ShapeRange.Count >= 1){     let olefTemp = Selection.ShapeRange.Item(1).OLEFormat     olefTemp.DisplayAsIcon = true     olefTemp.IconLabel = "My Icon" } }` |

#### **OLEFormat.IconName**

返回或设置存储 OLE 对象的图标的程序文件。**String** 类型，可读写。

**语法**

**express.IconName**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所选内容中第一个图形显示为图标，并将图标下面的文字设为该图标的文件名。*/ function test() { if(Selection.ShapeRange.Count >= 1){     let olefTemp = Selection.ShapeRange.Item(1).OLEFormat     olefTemp.DisplayAsIcon = true     olefTemp.IconLabel = olefTemp.IconName } }` |

#### **OLEFormat.IconPath**

返回存储 OLE 对象的图标的文件的路径。**String** 类型，只读。

**语法**

**express.IconPath**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示每个嵌入式 OLE 对象的路径，这些 OLE 对象在活动文档中显示为图标。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).Type == msoEmbeddedOLEObject){         if(ActiveDocument.Shapes.Item(i).OLEFormat.DisplayAsIcon == true){             MsgBox(shapeLoop.OLEFormat.IconPath)         }     } } }` |

#### **OLEFormat.Label**

返回用于标识源文件中被链接的部分的字符串。**String** 类型，只读。

**语法**

**express.Label**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

如果源文件是 ET 工作簿，并且 OLE 对象只包含工作表中少量的单元格，则 **Label** 属性可能返回“Workbook1!R3C1:R4C2”。

该属性仅对链接为 OLE 对象的图形、嵌入式图形或域有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动文档第一个域的标签。*/ MsgBox(ActiveDocument.Fields.Item(1).OLEFormat.Label)` |

#### **OLEFormat.Object**

返回一个 **Object** 类型的值，该值代表指定的 OLE 对象的最上层接口。

**语法**

**express.Object**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

使用该属性可以访问 ActiveX 控件或应用程序（创建 OLE 对象的应用程序）的属性和方法。为使该属性生效，OLE 对象必须支持“OLE 自动化”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第一个图形的值。为使本示例能够运行，第一个图形必须是 ActiveX 控件（例如，复选框或选项按钮）。*/ function test() { let myOA = ActiveDocument.Shapes.Item(1).OLEFormat myOA.Activate() myOA.Object.Value = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向活动文档中添加一个新的 ActiveX 控件，然后激活新添的选项按钮，并设置该按钮的部分属性。*/ function test() { let myOB = ActiveDocument.Shapes.AddOLEControl("Forms.OptionButton.1") let myOA = myOB.OLEFormat 	myOA.Activate()  let myObj = myOA.Object     myObj.Value = false     myObj.Caption = "My Caption"     myObj.AutoSize = true }` |

#### **OLEFormat.Parent**

返回一个 **Object** 类型值，该值代表指定 **OLEFormat** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **OLEFormat** 对象的变量。

#### **OLEFormat.PreserveFormattingOnUpdate**

如果为 **True**，则将 WPS 中的格式设置保存到链接的 OLE 对象中，例如链接到 ET 电子表格的表格。**Boolean** 类型，可读写。

**语法**

**express.PreserveFormattingOnUpdate**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

如果将 **PreserveFormattingOnUpdate** 设为 **True**，则更新对象时，保留在 WPS 中对该对象进行的格式更改。 WPS 仅更新链接对象中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*假定该文档中的第一个图形是链接的 OLE 对象，本示例保留当前文档中第一个图形的格式。*/ ActiveDocument.Shapes.Item(1).OLEFormat.PreserveFormattingOnUpdate = true ` |

#### **OLEFormat.ProgID**

返回指定 OLE 对象的程序标识符 (ProgID)?（程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或 PowerPoint.Slide），由 Windows 注册表用来唯一标识一个对象。）。只读 **String** 类型。

**语法**

**express.ProgID**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

**ProgID** 属性和 **ClassType** 属性（默认情况下）会返回相同的字符串。但是，您可以更改 DDE 链接的 **ClassType** 属性。

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

有关程序标识符的信息，请参阅 OLE 程序标识符。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例遍历活动文档中的所有浮动形状，并将所有链接的 ET 工作表设置为自动更新。*/ function test() { for(let i = 1; i <= ActiveDocument.Shapes.Count; i++){     if(ActiveDocument.Shapes.Item(i).Type == msoLinkedOLEObject){         if(ActiveDocument.Shapes.Item(i).OLEFormat.ProgID == "Excel.Sheet"){            ActiveDocument.Shapes.Item(i).LinkFormat.AutoUpdate = true         }     } } }` |

适用环境：web

适用平台：windows/linux