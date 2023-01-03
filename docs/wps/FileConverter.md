#### **FileConverter**



代表用于打开或保存文件的文件转换器。**FileConverter** 对象是 **FileConverters** 集合的一个成员。 **FileConverters** 集合包含所有已安装的用于打开和保存文件的文件转换器。

**说明**

使用 **FileConverters**(*Index*) 可返回一个 **FileConverter** 对象，其中 *Index* 是类别名称或索引号。以下示例显示与 ET 工作表转换器相关联的扩展名。

| 示例代码复制                                       |
| -------------------------------------------------- |
| `MsgBox(FileConverters.Item("MSBiff").Extensions)` |

索引号代表文件转换器在 **FileConverters** 集合中的位置。以下示例显示第一个文件转换器的格式名称。

| 示例代码复制                                |
| ------------------------------------------- |
| `MsgBox(FileConverters.Item(1).FormatName)` |

不能新建一个文件转换器或向 **FileConverters** 集合添加一个文件转换器。**FileConverter** 对象是在安装 WPS Office的过程中或通过安装补充文件转换器添加的。使用 **CanSave** 或 **CanOpen** 属性可确定 **FileConverter** 对象是否可用于打开或保存文档。

用于保存文档的文件转换器列在**“另存为”**对话框中。如果选中了**“选项”**对话框（**“工具”**菜单）中**“常规”**选项卡上的**“打开时确认转换”**复选框，则用于打开文档的文件转换器显示在一个对话框中。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanOpen**     | 如果指定的文件转换器可用来打开文件，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CanSave**     | 如果指定的文件转换器可用来保存文件，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ClassName**   | 返回用来标识文件转换器的唯一名字。**String** 类型，只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Extensions**  | 返回与指定 **FileConverter** 对象关联的文件扩展名。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormatName**  | 返回指定文件转换器的名称。**String** 类型，只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回指定对象的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OpenFormat**  | 返回指定文件转换器的文件格式。**Long** 类型，只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **FileConverter** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回指定对象的磁盘或 Web 路径。只读 **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveFormat**  | 返回指定文档或文件转换器的文件格式。只读 **Long** 类型。     |

**成员属性**

#### **FileConverter.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **FileConverter.CanOpen**

如果指定的文件转换器可用来打开文件，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.CanOpen**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

如果 **CanSave** 属性返回 **True**，则可以用指定的文件转换器保存（导出）文件。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断能否用第一个文件转换器打开文件。*/ function test() { if(FileConverters.Item(1).CanOpen == true) {     MsgBox(FileConverters.Item(1).FormatName + " can open files") } }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断能否用 WordPerfect6x 文件转换器打开文件。如果 CanOpen 属性返回 True，则打开名为“Test.wp”的文档。*/ function test() { if(FileConverters.Item("WordPerfect6x").CanOpen == true) {     Documents.Open("C:\\Test.wp",null,null,null,null,null,null,null,null,FileConverters.Item("WordPerfect6x").OpenFormat) } }` |

#### **FileConverter.CanSave**

如果指定的文件转换器可用来保存文件，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.CanSave**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

如果可以使用指定的文件转换器打开（导入）文件，则 **CanOpen** 属性返回 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断能否用 WordPerfect 转换器来保存文件。如果返回值为 True，则以 WordPerfect 6.x 格式保存活动文档。*/ function test() { if(Application.FileConverters.Item("WordPerfect6x").CanSave == true) {     let lngSaveFormat = Application.FileConverters.Item("WordPerfect6x").SaveFormat     ActiveDocument.SaveAs("C:\\Document.wp", lngSaveFormat) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，以表明 FileConverters 集合的第三个转换器能否用来保存文件。*/ function test() { if(FileConverters.Item(3).CanSave == true) {     MsgBox(FileConverters.Item(3).FormatName + " can save files") } else {     MsgBox(FileConverters.Item(3).FormatName + " cannot save files") } }` |

#### **FileConverter.ClassName**

返回用来标识文件转换器的唯一名字。**String** 类型，只读。

**语法**

**express.ClassName**

*express*   一个代表 **FileConverter** 对象的变量。

**示例**

------

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 FileConverters 集合中第一个转换器的类名和格式名。*/ MsgBox("ClassName= " + FileConverters.Item(1).ClassName + "\r" + "FormatName= " + FileConverters.Item(1).FormatName)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果 HTML 文件转换器有效，本示例将 HTML 设为默认保存格式。*/ function test() { for(let i=1;i <= FileConverters.Count;i++) {     if(FileConverters.Item(i).ClassName == "HTML") {         Application.DefaultSaveFormat = "HTML"     } } }` |

#### **FileConverter.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **FileConverter.Extensions**

返回与指定 **FileConverter** 对象关联的文件扩展名。**String** 类型，只读。

**语法**

**express.Extensions**

*express*   一个代表 **FileConverter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示第一个文件转换器的名称和文件扩展名。*/ function test() { let fcTemp = FileConverters.Item(1) MsgBox("The file name extensions for " + fcTemp.FormatName + " files are: " + fcTemp.Extensions) }` |

#### **FileConverter.FormatName**

返回指定文件转换器的名称。**String** 类型，只读。

**语法**

**express.FormatName**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

格式名称显示在**“文件”**菜单**“另存为”**对话框中的**“保存类型”**框中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 FileConverters 集合中的第一个文件转换器的格式名称。*/ MsgBox(FileConverters.Item(1).FormatName)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用 AvailableConv() 数组保存所有有效文件转换器的名称。*/ function test() { let AvailableConv = [] let intTemp = 0 for(let i=1;i <= FileConverters.Count;i++) {     AvailableConv[intTemp] = FileConverters.Item(i).FormatName     intTemp++ } }` |

#### **FileConverter.Name**

返回指定对象的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **FileConverter** 对象的变量。

#### **FileConverter.OpenFormat**

返回指定文件转换器的文件格式。**Long** 类型，只读。

**语法**

**express.OpenFormat**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

此属性可以是任何有效的 **WdOpenFormat** 常量，也可以是代表外部文件转换器的唯一数字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示可用来打开文档的转换器的唯一的格式值和格式名称。*/ function test() { for(let i=1;i <= FileConverters.Count;i++) {     if(FileConverters.Item(i).CanOpen == true) {         MsgBox(FileConverters.Item(i).OpenFormat + "\r" + FileConverters.Item(i).FormatName)     } } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使用 WordPerfect 6x 文件转换器打开名为“Data.wp”的文件。*/ Documents.Open("C:\\Data.wp",null,null,null,null,null,null,null,null,FileConverters.Item("WordPerfect6x").OpenFormat)` |

#### **FileConverter.Parent**

返回一个 **Object** 类型值，该值代表指定 **FileConverter** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **FileConverter** 对象的变量。

#### **FileConverter.Path**

返回指定对象的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

该路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符，使用 **Name** 属性可返回不带路径的文件名。您可以通过将 **Path**、**PathSeparator** 和 **Name** 属性连接在一起来创建文件转换器的完整名称。

#### **FileConverter.SaveFormat**

返回指定文档或文件转换器的文件格式。只读 **Long** 类型。

**语法**

**express.SaveFormat**

*express*   一个代表 **FileConverter** 对象的变量。

**说明**

该属性返回一个指定外部文件转换器的唯一数字或一个 **WdSaveFormat** 常量。使用 **SaveAs** 方法的 *FileFormat* 参数的 **SaveFormat** 属性值可将文档保存为没有相应的 **WdSaveFormat** 常量的文件格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一篇文档并在表格中列出可用于保存文档的转换器及其相应的 SaveFormat 值。*/ function FileConverterList() {     //Create a new document and set a tab stop     let docNew = Documents.Add()     docNew.Paragraphs.Format.TabStops.Add(InchesToPoints.Item(3))     //List all the converters in the FileConverters collection     let Con = docNew.Content     Con.InsertAfter("Name" + "\t" + "Number")     Con.InsertParagraphAfter()     for(let i=1;i <= FileConverters.Count;i++) {         if(FileConverters.Item(i).CanSave == true) {             Con.InsertAfter(FileConverters.Item(i).FormatName + "\t" + FileConverters.Item(i).SaveFormat)             Con.InsertParagraphAfter()         }     }     Con.ConvertToTable() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档保存为 WordPerfect 5.1 或 5.2 二级文件格式。*/ ActiveDocument.SaveAs(null,FileConverters.Item("WrdPrfctDat").SaveFormat)` |

适用环境：web

适用平台：windows/linux