#### **Hyperlink**



代表一个超链接。**Hyperlink** 对象是 **Hyperlinks** 集合的一个成员。

**说明**

使用 **Hyperlink** 属性可返回一个与形状相关联的 **Hyperlink** 对象（一个形状只能有一个超链接）。以下示例激活与活动文档中第一个形状相关联的超链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Shapes.Item(1).Hyperlink.Follow()` |

使用 **Hyperlinks**(*Index*) 可从文档、范围或所选内容返回单个 **Hyperlink** 对象，其中 *Index* 为索引号。以下示例激活所选内容中的第一个超链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     if(Application.Selection.Hyperlinks.Count >= 1){        Application.Selection.Hyperlinks.Item(1).Follow()     } }` |

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddToFavorites**    | 创建文档或超链接的快捷方式，并将其添加到“收藏夹”。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateNewDocument** | 新建一篇链接到指定超链接的文档。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除指定的超链接。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Follow**            | 显示一个与指定的 **Hyperlink** 对象相关的缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Address**           | 返回或设置指定超链接的地址（例如，文件名或 URL）。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailSubject**      | 返回或设置指定超链接的主题行的文本字符串。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ExtraInfoRequired** | 如果在解析指定的超链接时需要额外信息，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**              | 返回指定对象的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 返回一个 **Object** 类型值，该值代表指定 **Hyperlink** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**             | 返回一个 **Range** 对象，该对象代表超链接中包含的文档部分。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScreenTip**         | 返回或设置当鼠标指针置于指定超链接之上时显示的“屏幕提示”文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shape**             | 返回一个指定超链接或图表顶点的 **Shape** 对象。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SubAddress**        | 返回或设置指定超链接目标中命名的位置。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Target**            | 返回或设置在其中加载超链接的图文框或窗口的名称。可读写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextToDisplay**     | 返回或设置文档中指定超链接的可视文本。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 返回超链接的类型。只读 **MsoHyperlinkType** 类型。           |

**成员方法**

#### **Hyperlink.AddToFavorites**

创建文档或超链接的快捷方式，并将其添加到“收藏夹”。

**语法**

**express.AddToFavorites()**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的所有超链接创建快捷方式，并将其添加到“收藏夹”文件夹。*/ function test() {   for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){       Application.ActiveDocument.Hyperlinks.Item(i).AddToFavorites()   } }  /*本示例为 Sales.doc 创建快捷方式，并将其添加至“收藏夹”文件夹。如果 Sales.doc 还未打开，本示例将从 C:\Documents 文件夹打开该文档。*/ function test() {   let isOpen   for(let i = 1; i <= Application.Documents.Count; i++){       if(Application.Documents.Item(i).Name == "sales.doc"){          isOpen = true       }   }   if(isOpen != true){       Application.Documents.Open("C:\\Documents\\Sales.doc")       Application.Documents.Item("Sales.doc").AddToFavorites()   } }` |

#### **Hyperlink.CreateNewDocument**

新建一篇链接到指定超链接的文档。

**语法**

**express.CreateNewDocument(FileName, EditNow, Overwrite)**

*express*   一个代表 **Hyperlink** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*  | 必选          | **String**   | 指定文档的文件名。                                           |
| *EditNow*   | 必选          | **Boolean**  | 如果为 True，则立即在相关编辑环境中打开指定文档。默认值为 True。 |
| *Overwrite* | 必选          | **Boolean**  | 如果为 True，则覆盖相同文件夹中任何现有的同名文件。如果为 False，则保留所有现有的同名文件，并且参数 FileName 将指定一个新的文件名。默认值为 False。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例基于第一篇文档中新的超链接新建一篇文档，然后在 WPS 中加载该新文档以进行编辑。该文档名为“Overview.doc”，它将覆盖 \\Server1\Annual 文件夹中的任何同名文件。*/ function test() {   let objHyper = Application.Documents.Item(1).Hyperlinks.Add(Selection.Range, "\\Server1\\Annual\\Overview.doc")   objHyper.CreateNewDocument("\\Server1\\Annual\\Overview.doc", true, true) }` |

#### **Hyperlink.Delete**

删除指定的超链接。

**语法**

**express.Delete()**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Follow**

显示一个与指定的 **Hyperlink** 对象相关的缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。

**语法**

**express.Follow(NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)**

*express*   一个代表 **Hyperlink** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *NewWindow*  | 必选          | **Variant**  | 如果该属性值为 True，则在新窗口中显示目标文档。默认值为 False。 |
| *AddHistory* | 必选          | **Variant**  | 该参数为以后使用而保留。                                     |
| *ExtraInfo*  | 必选          | **Variant**  | 指定用于解析超链接的 HTTP 附加信息的字符串或字节数组。例如，您可以使用 ExtraInfo 来指定图像映射的坐标、窗体的内容或 FAT 文件的名称。根据 Method 的值来决定是发布还是附加该字符串。可使用 ExtraInfoRequired 属性确定是否需要额外信息。 |
| *Method*     | 必选          | **Variant**  | 指定 HTTP 附加信息的处理方式。可以是任何 MsoExtraInfoMethod 常量。 |
| *HeaderInfo* | 必选          | **Variant**  | 指定 HTTP 请求标题信息的字符串。默认值为一个空字符串。可以使用下列语法将几个标题行合并成一个字符串："string1" & vbCr & "string2"。指定的字符串自动转换为 ANSI 字符。注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标题字段。 |

**说明**

如果超链接使用文件协议，该方法不是下载而是打开该文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Home.doc 中的第一个超链接。*/ Application.Documents.Item("Home.doc").Hyperlinks.Item(1).Follow()  /*本示例首先插入指向 www.msn.com 的超链接，然后显示此超链接。*/ function test() {   Application.Selection.Collapse(wdCollapseEnd)   Application.Selection.InsertAfter("MSN ")   Application.Selection.Previous()   let hypTemp = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://www.msn.com")       hypTemp.Follow(false, true) }` |

**成员属性**

#### **Hyperlink.Address**

返回或设置指定超链接的地址（例如，文件名或 URL）。**String** 类型，可读写。

**语法**

**express.Address**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果不存在与某对象关联的超链接，则设置 **Address** 属性将返回一个错误。在此情况下，请将**Add**方法用于 **Hyperlinks** 集合来添加超链接。以下示例说明了具体的操作方法。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Hyperlinks.Add(Selection.Range,"http://www.wps.cn")` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能是：将超链接添加到活动文档的选定内容，设置地址然后在消息框中显示该地址。*/ let aHLink = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://forms") alert("The hyperlink goes to " + aHLink.Address)  /*如果活动文档包含超链接，则本示例会在文档的末尾插入超链接目标的列表。*/ let myRange = Application.ActiveDocument.Range(ActiveDocument.Content.End - 1) let Count = 0 for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){     Count = Count + 1     myRange.InsertAfter("Hyperlink #" + Count + "\t")     myRange.InsertAfter(Application.ActiveDocument.Hyperlinks.Item(i).Address) 	myRange.InsertParagraphAfter() }` |

#### **Hyperlink.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Hyperlink.EmailSubject**

返回或设置指定超链接的主题行的文本字符串。**String** 类型，可读写。

**语法**

**express.EmailSubject**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

该主题行附加到超链接的 Internet 地址或 URL 中。该属性通常用于电子邮件超链接。其值优先于在相同 **Hyperlink** 对象的 Address属性中指定的其他任何电子邮件主题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：检查活动文档中的电子邮件超链接。如果发现某个超链接有空主题行，则为其添加主题“NewProducts”。*/ function test() {   for(let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++){       if(Application.ActiveDocument.Hyperlinks.Item(i).Address.indexOf("mailto*") == 0 && Application.ActiveDocument.Hyperlinks.Item(i).Address == ActiveDocument.Hyperlinks.Item(i).EmailSubject){           Application.ActiveDocument.Hyperlinks.Item(i).EmailSubject = "NewProducts"       }   } }` |

#### **Hyperlink.ExtraInfoRequired**

如果在解析指定的超链接时需要额外信息，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.ExtraInfoRequired**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

可以将 *ExtraInfo* 参数与**Follow** 或 **FollowHyperlink**方法结合使用，以便指定其他信息。例如，可以使用 *ExtraInfo* 指定图像映射的坐标、窗体的内容或 FAT 文件名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先插入指向 www.msn.com 的超链接，如果不需要额外信息，则跟踪此超链接。*/ function test() {   Application.Selection.Collapse(wdCollapseEnd)   Application.Selection.InsertAfter("MSN ")   Application.Selection.Previous()      let hypTemp = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "http://www.msn.com")   if(hypTemp.ExtraInfoRequired == false){      hypTemp.Follow()   } }` |

#### **Hyperlink.Name**

返回指定对象的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Parent**

返回一个 **Object** 类型值，该值代表指定 **Hyperlink** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Range**

返回一个 **Range** 对象，该对象代表超链接中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.ScreenTip**

返回或设置当鼠标指针置于指定超链接之上时显示的“屏幕提示”文字。**String** 类型，可读写。

**语法**

**express.ScreenTip**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第一个超链接的“屏幕提示”文字。*/ Application.ActiveDocument.Hyperlinks.Item(1).ScreenTip = "Home"` |

#### **Hyperlink.Shape**

返回一个指定超链接或图表顶点的 **Shape** 对象。

**语法**

**express.Shape**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果超链接不是由形状表示的，则会产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例改变形状的填充颜色，该形状代表活动文档中的第一个超链接。为使本示例能运行，超链接必须由形状表示。*/ Application.ActiveDocument.Hyperlinks.Item(1).Shape.Fill.ForeColor.RGB = (255, 255, 0)` |

#### **Hyperlink.SubAddress**

返回或设置指定超链接目标中命名的位置。**String** 类型，可读写。

**语法**

**express.SubAddress**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

命名的位置可以是 WPS 文档中的书签、ET 工作表中命名的单元格或单元格引用、Access 数据库中命名的对象，或者是 WPP 演示文稿中幻灯片的序号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本例显示所选超链接的子地址。*/ function test() {   if(Application.Selection.Range.Hyperlinks.Count >= 1){       alert(Application.Selection.Range.Hyperlinks.Item(1).SubAddress)   } }  /*本例为活动文档中的选定内容添加超链接，设置超链接的目标和子地址，然后将其显示在消息框中。*/ function test() {   let SCut = Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "C:\\My Documents\\Other.doc", "temp")   alert("The hyperlink goes to " + SCut.SubAddress) }` |

#### **Hyperlink.Target**

返回或设置在其中加载超链接的图文框或窗口的名称。可读写 **String** 类型。

**语法**

**express.Target**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置在一个新的浏览器窗口中打开指定的超链接。*/ Application.ActiveDocument.Hyperlinks.Item(1).Target = "_blank"  /*以下示例设置在名为“left”的图文框中打开指定的超链接。*/ Application.ActiveDocument.Hyperlinks.Item(1).Target = "left"` |

#### **Hyperlink.TextToDisplay**

返回或设置文档中指定超链接的可视文本。**String** 类型，可读写。

**语法**

**express.TextToDisplay**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第一个超链接的显示文本。*/ Application.ActiveDocument.Hyperlinks.Item(1).TextToDisplay = "Follow this link for more information..."` |

#### **Hyperlink.Type**

返回超链接的类型。只读 **MsoHyperlinkType** 类型。

**语法**

**express.Type**

*express*   一个代表 **Hyperlink** 对象的变量。

适用环境：web

适用平台：windows/linux