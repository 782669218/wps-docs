**Shapes**



指定的工作表上的所有 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象的集合。

**说明**

每个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象都代表绘图层中的一个对象，如自选图形、任意多边形、OLE 对象或图片。

**注释：**如果您想处理文档中的一部分形状（例如，只针对文档中的自选图形或只对选定的形状进行一些操作），则必须构造一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)集合，其中包含您要处理的形状。

## 示例

使用 **Shapes** 属性可返回 **Shapes**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合。下例选定 myDocument 上的所有形状。

**注释：**如果您要同时对工作表上的所有形状进行操作（例如删除或设置属性），请选定所有形状，然后对选定区域使用 **ShapeRange** 属性，以创建一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象，该对象包含工作表上的所有形状，然后对 **ShapeRange** 对象应用相应的属性或方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.SelectAll() }` |

 

**使用 Shapes(index)（其中 index 是形状的名称或索引号）可返回一个 Shape 对象。下例设置 myDocument 上形状一的预设阴影的填充。**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Item(1).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass) }` |

**使用 Shapes.Range(index)（其中 index 是形状的名称或索引号，或是它们的一个数组）可返回一个 ShapeRange 集合，该集合代表 Shapes 集合的一个子集。下例设置 myDocument 上形状一和三的填充图案。**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick) }` |

工作表上的 ActiveX 控件有两个名称：一个是包含该控件的形状的名称，查看工作表时可在**“名称”**框中看到它；另一个是控件的代码名称，可以在**“属性”**窗口中**“(名称)”**右边的单元格中看到它。在您首次向工作表添加控件时，形状名称和代码名称是一致的。但是，如果您更改这两个名称中的任意一个，另一个不会随之自动更改。

在控件事件过程名称中使用的是控件的代码名称，但是，当您从工作表的 **Shapes** 或 [**OLEObjects** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/OLEObjects/OLEObjects%20.htm#jsObject_OLEObjects)集合中返回控件时，必须使用形状名称而不是代码名称，以便按名称引用控件。例如，假定您给工作表添加一个复选框，而默认的形状名称和代码名称都是 CheckBox1。如果通过在“属性”窗口的“(名称)”旁边键入“chkFinished”而更改了控件代码名称，则在事件过程名称中必须使用 chkFinished，但是您仍然需要使用 CheckBox1 从 **Shapes** 或 [**OLEObject** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/OLEObject/OLEObject%20.htm#jsObject_OLEObject)集合中返回控件，如下例所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 }` |

**方法**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddOLEObject** | 创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。返回值类型为Shape。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddPicture**   | 现有文件创建图片。返回一个代表新图片的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。返回值类型为Shape。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddShape**     | 返回一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，该对象表示工作表中的新自选形状。返回值类型为Shape。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**         | 从集合中返回一个对象。返回值：包含在集合中的一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**        | 返回一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象，它代表 **Shapes** 集合中形状的子集。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectAll**    | 选择指定的 **Shapes** 集合中的所有形状。                     |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |

**成员方法**

#### **Shapes.AddOLEObject**

创建 OLE 对象。返回一个代表新 OLE 对象的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

返回值类型为Shape。

**语法**

**express.AddOLEObject(ClassType, Filename, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | （必须指定 ClassType 或 FileName）。该字符串包含要创建的对象的程序标识符。如果指定了 ClassType 参数，则忽略 FileName 和 Link。 |
| *Filename*      | 可选          | **Variant**  | 创建对象所用的文件名称。如果未指定路径，则使用当前的工作文件夹。必须为对象指定 ClassType 或 FileName 参数，但不能同时指定。 |
| *Link*          | 可选          | **Variant**  | 如果为 True，则将 OLE 对象链接到创建该对象所使用的文件；如果为 False，则 OLE 对象成为该文件的独立副本。如果为 ClassType 指定值，则该参数必须为 False。默认值为 False。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果为 True，则将 OLE 对象显示为图标。默认值为 False。       |
| *IconFileName*  | 可选          | **Variant**  | 包含将要显示的图标的文件。                                   |
| *IconIndex*     | 可选          | **Variant**  | IconFileName 内的图标索引。指定文件中图标的顺序与图标在“更改图标”对话框中出现的顺序相对应（选中“显示为图标”复选框后，可从“对象”对话框访问该对话框）。文件中的第一个图标的索引号为 0（零）。如果 IconFileName 中不存在给定索引号的图标，则使用索引号为 1 的图标（即文件中的第二个图标）。默认值为 0（零）。 |
| *IconLabel*     | 可选          | **Variant**  | 显示在图标下面的标签（题注）。                               |
| *Left*          | 可选          | **Variant**  | 新对象的左上角相对于文档左上角的位置（以磅为单位）。默认值为 0（零）。 |
| *Top*           | 可选          | **Variant**  | 新对象的左上角相对于文档左上角的位置（以磅为单位）。默认值为 0（零）。 |
| *Width*         | 可选          | **Variant**  | 以磅为单位给出 OLE 对象的初始尺寸。                          |
| *Height*        | 可选          | **Variant**  | 以磅为单位给出 OLE 对象的初始尺寸。                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加链接的 WPS 文档。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.AddOLEObject(null,"c:\\my documents\\testing.doc",true,null,null,null,null,100,100,200,300) }  /*本示例向 myDocument 中添加新的命令按钮。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.AddOLEObject("Forms.CommandButton.1",null,null,null,null,null,null,100,100,100,200) }` |

#### **Shapes.AddPicture**

现有文件创建图片。返回一个代表新图片的 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

返回值类型为Shape。

**语法**

**express.AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**    | **说明**                                         |
| ------------------ | ------------- | --------------- | ------------------------------------------------ |
| *Filename*         | 必选          | **String**      | 要在其中创建 OLE 对象的文件。                    |
| *LinkToFile*       | 必选          | **MsoTriState** | 要链接至的文件。                                 |
| *SaveWithDocument* | 必选          | **MsoTriState** | 将图片与文档一起保存。                           |
| *Left*             | 必选          | **Single**      | 图片左上角相对于文档左上角的位置（以磅为单位）。 |
| *Top*              | 必选          | **Single**      | 图片左上角相对于文档顶部的位置（以磅为单位）。   |
| *Width*            | 必选          | **Single**      | 图片的宽度（以磅为单位）。                       |
| *Height*           | 必选          | **Single**      | 图片的高度（以磅为单位）。                       |

**说明**

| **MsoTriState** 可以是下列 [**MsoTriState** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoTriState%20%E6%9E%9A%E4%B8%BE.html)常量之一。 |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 使图片成为其源文件的独立副本。                  |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 建立图片与其源文件之间的链接。                   |

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 在文档中只存储链接信息。                        |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 将链接图片与该图片插入的文档一起保存。如果 LinkToFile 为 **msoFalse**，则该参数必须为 **msoTrue**。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加由文件“Music.bmp”创建的图片。插入的图片链接到创建该图片的文件，并与 myDocument 一起保存。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.AddPicture("c:\\microsoft office\\clipart\\music.bmp", true, true, 100, 100, 70, 70) }` |

#### **Shapes.AddShape**

返回一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象，该对象表示工作表中的新自选形状。

返回值类型为Shape。

**语法**

**express.AddShape(Type, Left, Top, Width, Height)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**         | **说明**                                                   |
| -------- | ------------- | -------------------- | ---------------------------------------------------------- |
| *Type*   | 必选          | **MsoAutoShapeType** | 指定要创建的自选形状的类型。                               |
| *Left*   | 必选          | **Single**           | 自选形状边框的左上角相对于文档左上角的位置（以磅为单位）。 |
| *Top*    | 必选          | **Single**           | 自选形状边框的左上角相对于文档左上角的位置（以磅为单位）。 |
| *Width*  | 必选          | **Single**           | 自选形状边框的宽度（以磅为单位）。                         |
| *Height* | 必选          | **Single**           | 自选形状边框的高度（以磅为单位）。                         |

**说明**

要更改已添加的自选形状的类型，请设置 **AutoShapeType** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 myDocument 中添加矩形。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200) }` |

#### **Shapes.Item**

从集合中返回一个对象。

返回值：包含在集合中的一个 **Shape**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Shapes 集合中第二个形状的 OnAction 属性进行设置。如果变量 ss 代表的不是 Shapes 对象，则本示例无效。*/ function test() {     let ss = Application.Worksheets.Item(1).Shapes     ss.Item(2).OnAction = "ShapeAction" }` |

#### **Shapes.Range**

返回一个 **ShapeRange**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象，它代表 **Shapes** 集合中形状的子集。

**语法**

**express.Range(Index)**

*express*   一个代表 **Shapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | t 包含在该区域中的各单个形状。可以是指定形状索引号的整数、指定形状名称的字符串，也可以是包含整数或字符串的数组。 |

**说明**

虽然使用 **Range** 属性可返回任意数量的形状，但如果要返回集合中单个成员时，用 [**Item** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#Shapes.Item)方法更加简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。例如，以下指令返回用名称指定的两个形状。

`Dim arShapes() As Variant Dim objRange As Object arShapes = Array("Oval 4", "Rectangle 5") Set objRange = ActiveSheet.Shapes.Range(arShapes)`

在 ET 中，不能用此属性返回包含工作表上的所有 [**Shape** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shape/Shape%20.htm#jsObject_Shape)对象的 [**ShapeRange** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/ShapeRange/ShapeRange%20.htm#jsObject_ShapeRange)对象。如果要达到该目的，可用下列代码：

`Worksheets(1).Shapes.SelectAll ' select all shapes set sr = Selection.ShapeRange ' create ShapeRange`

#### **Shapes.SelectAll**

选择指定的 **Shapes** 集合中的所有形状。

**语法**

**express.SelectAll()**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择 myDocument 中的所有形状，并创建包含所有这些形状的 ShapeRange 集合。*/ function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.SelectAll()      let sr = Application.Selection.ShapeRange }` |

**成员属性**

#### **Shapes.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Shapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET"){         alert("This is an ET Application object.")     }     else{         alert("This is not an ET Application object.")     } }` |

#### **Shapes.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Shapes** 对象的变量。

适用环境：web

适用平台：windows/linux