**Workbook**



代表一个 ET 工作簿。

**说明**

**Workbook** 对象是 [Workbooks ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)集合的成员。**Workbooks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)集合包含 ET 中当前打开的所有 **Workbook** 对象。

## 示例

使用 **Workbooks**(*index*)（其中 *index* 是工作簿名称或索引号）可返回一个 **Workbook**对象。下例激活工作簿一。

| 示例代码复制                               |
| ------------------------------------------ |
| `Application.Workbooks.Item(1).Activate()` |

编号指示创建或打开工作簿的顺序。`Workbooks(1)` 是创建的第一个工作簿，而 `Workbooks(Workbooks.Count)`Workbooks 是最后一个。激活某工作簿并不更改其索引号。所有工作簿均包括在索引计数中，即便是隐藏工作簿也是如此。

[Name ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#Workbook.Name)属性返回工作簿名称。您不能通过使用此属性来设置该名称；如果您需要更改该名称，请使用 [SaveAs ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#Workbook.SaveAs)方法，将该工作簿保存为其他名称。下例激活名为“Cogs.xls”的工作簿（该工作簿必须已经在 ET 中打开）中的 Sheet1。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Workbooks.Item("Cogs.xls").Worksheets.Item("Sheet1").Activate()` |

**ActiveWorkbook**属性返回当前处于活动状态的工作簿。下例设置活动工作簿作者的名称。

| 示例代码复制                                       |
| -------------------------------------------------- |
| `Application.ActiveWorkbook.Author = "Jean Selva"` |

**方法**

|                                                              | 名称           | 说明                                                       |
| ------------------------------------------------------------ | -------------- | ---------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**   | 激活与工作簿相关的第一个窗口。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**      | 关闭对象。                                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**  | 新建一个窗口或者创建指定窗口的副本。返回值类型为Window。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**       | 保存对指定工作簿所做的更改。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAs**     | 在另一不同文件中保存对工作簿所做的更改。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveCopyAs** | 将指定工作簿的副本保存到文件，但不修改内存中的打开工作簿。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Activesheet** | 返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **Nothing**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FileFormat**  | 返回工作簿的文件格式和/或类型。[XlFileFormat ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlFileFormat%20%E6%9E%9A%E4%B8%BE.html)类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**    | 返回对象的名称（以字符串表示），包括其磁盘路径。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasPassword** | 如果指定工作簿有密码保护，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Names**       | 返回一个 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，它代表指定工作簿的所有名称（包括所有指定工作表的名称）。**Names** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**    | 如果对象以只读方式打开，则返回 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Saved**       | 如果指定工作簿从上次保存至今未发生过更改，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sheets**      | 返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中所有工作表。**Sheets** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**     | 返回一个 [Windows ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Windows/Windows%20.htm#jsObject_Windows)集合，它代表指定工作簿中所有窗口。**Windows** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Worksheets**  | 返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中所有工作表。**Sheets** 对象，只读。 |

**成员方法**

#### **Workbook.Activate**

激活与工作簿相关的第一个窗口。

**语法**

**express.Activate()**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

这样不会运行可能附加在工作簿上的 Auto_Activate 或 Auto_Deactivate 宏（使用**RunAutoMacros**方法可运行这些宏）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活工作簿 Book4.xls。如果工作簿 Book4.xls 有若干窗口，本示例激活第一个，即 Book4.xls:1。*/ Application.Workbooks.Item("BOOK4.XLS").Activate()` |

#### **Workbook.Close**

关闭对象。

**语法**

**express.Close(SaveChanges, Filename, RouteWorkbook)**

*express*   一个代表 **Workbook** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges*   | 可选          | **Variant**  | 如果工作簿中没有改动，则忽略此参数。如果工作簿中有改动但工作簿显示在其他打开的窗口中，则忽略此参数。如果工作簿中有改动且工作簿未显示在任何其他打开的窗口中，则由此参数指定是否应保存更改。如果设为 True，则保存对工作簿所做的更改。如果工作簿尚未命名，则使用 FileName。如果省略 Filename，则要求用户提供文件名。 |
| *Filename*      | 可选          | **Variant**  | 以此文件名保存所做的更改。                                   |
| *RouteWorkbook* | 可选          | **Variant**  | 如果工作簿不需要传送给下一个收件人（没有传送名单或已经传送），则忽略此参数。否则，ET 根据此参数的值传送工作簿。如果设为 True，则将工作簿传送给下一个收件人。如果设为 False，则不发送工作簿。如果忽略，则要求用户确认是否发送工作簿。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例关闭 Book1.xls，并放弃所有对此工作簿的更改。*/ Application.Workbooks.Item("BOOK1.XLS").Close(false)` |

#### **Workbook.NewWindow**

新建一个窗口或者创建指定窗口的副本。

返回值类型为Window。

**语法**

**express.NewWindow()**

*express*   一个代表 **Workbook** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为当前活动工作簿新建一个窗口。*/ Application.ActiveWorkbook.NewWindow()` |

#### **Workbook.Save**

保存对指定工作簿所做的更改。

**语法**

**express.Save()**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

要打开工作簿文件，请使用**Open**方法。

要将工作簿标记为已保存，但不将其写入磁盘，请将它的**Saved** 属性设置为 **True**。

首次保存工作簿时，请使用 [SaveAs ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#Workbook.SaveAs)方法指定文件名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保存活动工作簿。*/ Appliction.ActiveWorkbook.Save()  /*本示例保存所有打开的工作簿，然后关闭 ET。*/ function test() {     for (let w = 1; w <= Application.Workbooks.Count; w++) {         Application.Workbooks.Item(w).Save()     }     Application.Quit() }` |

#### **Workbook.SaveAs**

在另一不同文件中保存对工作簿所做的更改。

**语法**

**express.SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)**

*express*   一个代表 **Workbook** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型**                 | **说明**                                                     |
| --------------------- | ------------- | ---------------------------- | ------------------------------------------------------------ |
| *Filename*            | 可选          | **Variant**                  | 一个表示要保存文件的文件名的字符串。可包含完整路径，如果不指定路径，ET 将文件保存到当前文件夹中。 |
| *FileFormat*          | 可选          | **Variant**                  | 保存文件时使用的文件格式。要查看有效的选项列表，请参阅 XlFileFormat 枚举。对于现有文件，默认采用上一次指定的文件格式；对于新文件，默认采用当前所用 ET 版本的格式。 |
| *Password*            | 可选          | **Variant**                  | 它是一个区分大小写的字符串（最长不超过 15 个字符），用于指定文件的保护密码。 |
| *WriteResPassword*    | 可选          | **Variant**                  | 一个表示文件写保护密码的字符串。如果文件保存时带有密码，但打开文件时不输入密码，则该文件以只读方式打开。 |
| *ReadOnlyRecommended* | 可选          | **Variant**                  | 如果为 True，则在打开文件时显示一条消息，提示该文件以只读方式打开。 |
| *CreateBackup*        | 可选          | **Variant**                  | 如果为 True，则创建备份文件。                                |
| *AccessMode*          | 可选          | **XlSaveAsAccessMode**       | 工作簿的访问模式。                                           |
| *ConflictResolution*  | 可选          | **XlSaveConflictResolution** | 一个 XlSaveConflictResolution 值，它确定该方法在保存工作簿时如何解决冲突。如果设为 xlUserResolution，则显示冲突解决对话框。如果设为 xlLocalSessionChanges，则自动接受本地用户的更改。如果设为 xlOtherSessionChanges，则自动接受来自其他会话的更改（而不是本地用户的更改）。如果省略此参数，则显示冲突处理对话框。 |
| *AddToMru*            | 可选          | **Variant**                  | 如果为 True，则将该工作簿添加到最近使用的文件列表中。默认值为 False。 |
| *TextCodepage*        | 可选          | **Variant**                  | ET中对于所有语言都忽略此参数。 注释:当 ET 将工作簿保存为某种 CSV 或文本格式（使用 FileFormat 参数指定）时， ET 使用对应于当前计算机上使用的系统区域设置语言的代码页。在“控制面板”中单击“区域和语言”，再单击“位置”选项卡，在“当前位置”下可获得此系统设置。 |
| *TextVisualLayout*    | 可选          | **Variant**                  | ET中对于所有语言都忽略此参数。 注释:当 ET 将工作簿保存为某种 CSV 或文本格式（使用 FileFormat 参数指定）时，它按逻辑布局保存这些格式。如果文件中左至右 (LTR) 文本嵌在右至左 (RTL) 文本中，或者相反，那么逻辑布局将把文件的内容，按照文件中所有语言的正确阅读顺序保存，而不考虑方向。当应用程序打开文件时，每串 LTR 或 RTL 字符将根据代码页中的字符值范围，按照正确的方向呈现。（除非用来打开文件的应用程序是为显示文件的确切内存布局而设计的应用程序，如调试器或编辑器）。 |
| *Local*               | 可选          | **Variant**                  | 如果为 True，则以 ET（包括控制面板设置）的语言保存文件。如果为 False（默认值），则以 示例代码 (VBA) 的语言保存文件。VBA 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的国际化 XL5/95 VBA 项目。 |

**说明**

请使用同时包含大小写字母、数字和符号的强密码。弱密码不混合使用这些元素。强密码：Y6dh!et5。弱密码：House27。请使用您可以记住的强密码，这样就不必将它写下来。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一个工作簿，提示用户输入文件名，然后保存该工作簿。*/ function test() {     let NewBook = Application.Workbooks.Add()     do {         fName = Application.GetSaveAsFilename()     }     while(fName == false)     NewBook.SaveAs(fName) }` |

#### **Workbook.SaveCopyAs**

将指定工作簿的副本保存到文件，但不修改内存中的打开工作簿。

**语法**

**express.SaveCopyAs(Filename)**

*express*   一个代表 **Workbook** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**           |
| ---------- | ------------- | ------------ | ------------------ |
| *Filename* | 可选          | **Variant**  | 指定副本的文件名。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保存活动工作簿的副本。*/ Application.ActiveWorkbook.SaveCopyAs("C:\\TEMP\\XXXX.XLS")` |

**成员属性**

#### **Workbook.Activesheet**

返回一个对象，它代表活动工作簿中或指定的窗口或工作簿中的活动工作表（最上面的工作表）。如果没有活动的工作表，则返回 **Nothing**。

**语法**

**express.Activesheet**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

如果不指定对象识别符，则此属性返回活动工作簿中的活动工作表。

如果某个工作簿出现在若干个窗口中，那么该工作簿的 **ActiveSheet** 属性在不同窗口中可能不同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示活动工作表的名称。*/ alert("The name of the active sheet is " + Application.ActiveSheet.Name)` |

#### **Workbook.Application**

如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Workbook** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if(Application.ActiveWorkbook.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Workbook.FileFormat**

返回工作簿的文件格式和/或类型。[XlFileFormat ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlFileFormat%20%E6%9E%9A%E4%B8%BE.html)类型，只读。

**语法**

**express.FileFormat**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

这些常量中的某些可能不可用，这取决于选择或安装的语言支持（例如，美国英语）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查当前工作簿文件格式是否为 Excel9795 格式，如果是，则按常规文件格式保存该工作簿。*/ function test() {     if(Application.ActiveWorkbook.FileFormat == xlExcel9795) {         Application.ActiveWorkbook.SaveAs(null, xlExcel12)     }   }` |

#### **Workbook.FullName**

返回对象的名称（以字符串表示），包括其磁盘路径。**String** 型，只读。

**语法**

**express.FullName**

*express*   一个代表 **Workbook** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示当前工作簿的路径及文件名（假定尚未保存此工作簿）。*/ alert(Application.ActiveWorkbook.FullName)` |

#### **Workbook.HasPassword**

如果指定工作簿有密码保护，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasPassword**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

可以使用 [SaveAs](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#Workbook.SaveAs) 方法为工作簿分配一个保护密码。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动工作簿是否有密码保护，如果有则显示一条消息。*/ function test() {     if(Application.ActiveWorkbook.HasPassword == true) {         alert("Remember to obtain the workbook password" + "\r" +             " from the Network Administrator.")     } }` |

#### **Workbook.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Workbook** 对象的变量。

#### **Workbook.Names**

返回一个 **Names**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，它代表指定工作簿的所有名称（包括所有指定工作表的名称）。**Names** 对象，只读。

**语法**

**express.Names**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例是将 Sheet1 中的 A1 单元格的名称定义为“myName”。*/ Application.ActiveWorkbook.Names.Add("myName",null,null,null,null,null,null,null,null,"=Sheet1!R1C1")` |

#### **Workbook.Path**

返回一个 **String** 值，它代表应用程序的完整路径，不包括末尾的分隔符和应用程序名称。

**语法**

**express.Path**

*express*   一个代表 **Workbook** 对象的变量。

#### **Workbook.ReadOnly**

如果对象以只读方式打开，则返回 **True**。**Boolean** 类型，只读。

**语法**

**express.ReadOnly**

*express*   一个代表 **Workbook** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前工作簿是只读的，则此示例将此工作簿另存为 Newfile.xls。*/ function test() {     if(Application.ActiveWorkbook.ReadOnly == true) {         Application.ActiveWorkbook.SaveAs("NEWFILE.XLS")     } }` |

#### **Workbook.Saved**

如果指定工作簿从上次保存至今未发生过更改，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Saved**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

如果某个工作簿从未保存过，则其  **Path**属性返回一个空字符串 ("")。

如果要关闭某个已更改的工作簿，但又不想保存它或者不想出现保存提示，则可将此属性设为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动工作簿是否有未保存的更改，如果有，则显示一条消息。*/ function test() {     if(Application.ActiveWorkbook.Saved == false) {         alert("This workbook contains unsaved changes.")     } }` |

#### **Workbook.Sheets**

返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中所有工作表。**Sheets** 对象，只读。

**语法**

**express.Sheets**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Sheets`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例新建一张工作表，然后在第一列中列出活动工作簿中的工作表名称。*/ function test() {     let newSheet = Application.Sheets.Add(null,null,null,xlWorksheet)     for(let i = 1; i <= Application.Sheets.Count; i++) {         newSheet.Cells.Item(i, 1).Value2 = Application.Sheets.Item(i).Name     } }` |

#### **Workbook.Windows**

返回一个 [Windows ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Windows/Windows%20.htm#jsObject_Windows)集合，它代表指定工作簿中所有窗口。**Windows** 对象，只读。

**语法**

**express.Windows**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `Application.Windows`。

此属性返回的集合中既包括可见窗口，也包括隐藏窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将活动工作簿的窗口一命名为“Consolidated Balance Sheet”。此名称将被用作对 Windows 集合的索引。*/ function test() {     Application.ActiveWorkbook.Windows.Item(1).Caption = "Consolidated Balance Sheet"     Application.ActiveWorkbook.Windows.Item("Consolidated Balance Sheet").ActiveSheet.Calculate() }` |

#### **Workbook.Worksheets**

返回一个 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合，它代表指定工作簿中所有工作表。**Sheets** 对象，只读。

**语法**

**express.Worksheets**

*express*   一个代表 **Workbook** 对象的变量。

**说明**

在不使用对象识别符的情况下，使用此属性将返回活动工作簿中所有的工作表。

此属性不返回宏表；使用 **Excel4MacroSheets** 属性或 **Excel4IntlMacroSheets** 属性可返回这些表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示活动工作簿中 sheet1 上单元格 A1 中的值。*/ alert(Application.Worksheets.Item("Sheet1").Range("A1").Value2)  /*此示例显示活动工作簿中每个工作表的名称。*/ function test() {     for(let i = 1; i <= Application.Worksheets.Count; i++) {         alert(Application.Worksheets.Item(i).Name)     } }  /*此示例向活动工作簿添加新工作表，并设置该工作表的名称。*/ function test() {     let newSheet = Application.Worksheets.Add()     newSheet.Name = "current Budget" }` |

适用环境：web

适用平台：windows/linux