**Hyperlink**



代表一个超链接。

**说明**

**Hyperlink** 对象是 **Hyperlinks** 集合的成员。

使用 **Hyperlink** 属性可返回某个形状的超链接（一个形状只能有一个超链接）。下例激活形状一的超链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)` |

区域或工作表可以有多个超链接。使用 **Hyperlinks**(*index*) 可返回单个 **Hyperlink** 对象，其中 *index* 是超链接编号。下例激活 A1:B2 区域中的第二个超链接。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("A1:B2").Hyperlinks.Item(2).Follow()` |

**方法**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddToFavorites**    | 将工作簿或超链接的快捷方式添加到“收藏夹”文件夹。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateNewDocument** | 新建一篇链接到指定超链接的文档。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**            | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Follow**            | 如果已经下载指定文档，则显示缓冲区中的该文档。否则，本方法对指定超链接进行处理以下载目标文档，然后将该文档在适当的应用程序中显示出来。 |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Address**       | 返回或设置一个 **String** 型，它代表目标文档的地址。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailSubject**  | 返回或设置指定超链接的电子邮件主题行的文本字符串。主题行是添加到超链接地址上的。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**          | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**         | 返回一个 **Range** 对象，它代表附加的指定超链接的区域。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScreenTip**     | 返回或设置指定超链接的屏幕提示文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shape**         | 返回一个 **Shape** 对象，它代表附加到指定超链接的形状。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SubAddress**    | 返回或设置文档中与指定超链接相关联的位置。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextToDisplay** | 返回或设置要为指定超链接显示的文本。默认值为超链接的地址。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**          | 返回一个包含 **MsoHyperlinkType** 常量的 **Long** 值，它代表 HTML 框架的位置。 |

**成员方法**

#### **Hyperlink.AddToFavorites**

将工作簿或超链接的快捷方式添加到“收藏夹”文件夹。

**语法**

**express.AddToFavorites()**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例将活动工作簿的快捷方式添加到“收藏夹”文件夹*/ Application.ActiveWorkbook.AddToFavorites()` |

#### **Hyperlink.CreateNewDocument**

新建一篇链接到指定超链接的文档。

**语法**

**express.CreateNewDocument(Filename, EditNow, Overwrite)**

*express*   一个代表 **Hyperlink** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Filename*  | 必选          | **String**   | 指定文档的文件名。                                           |
| *EditNow*   | 必选          | **Boolean**  | 如果该值为 True，则立即在关联的编辑环境中打开指定文档。默认值为 True。 |
| *Overwrite* | 必选          | **Boolean**  | 如果该值为 True，则覆盖相同文件夹中任何现有的同名文件。如果该值为 False，则保留任何现有的同名文件，并且参数 Filename 将指定一个新的文件名。默认值为 False。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在第一张工作表中新建一个基于新的超链接的文档，然后将该文档加载到 ET 中进行编辑。该文档名为“Report.xls”，它将覆盖“\\Server1\Annual”文件夹中的所有同名文件*/ function test() {   let sheet2 = Application.Worksheets.Item(1)   let objHyper = sheet2.Hyperlinks.Add(sheet2.Range("A10"), "\\\Server1\\Annual\\Report.xls")   objHyper.CreateNewDocument("\\\Server1\\Annual\\Report.xls", true, true) }` |

#### **Hyperlink.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Follow**

如果已经下载指定文档，则显示缓冲区中的该文档。否则，本方法对指定超链接进行处理以下载目标文档，然后将该文档在适当的应用程序中显示出来。

**语法**

**express.Follow(NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)**

*express*   一个代表 **Hyperlink** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *NewWindow*  | 可选          | **Variant**  | 如果为 True，则在新窗口中显示目标应用程序。默认值为 False。  |
| *AddHistory* | 可选          | **Variant**  | 未使用。保留供将来使用。                                     |
| *ExtraInfo*  | 可选          | **Variant**  | 指定解析超链接时要使用的 HTTP 附加信息的 String 或字节数组。例如，可以使用 ExtraInfo 指定图像的坐标、窗体的内容或 FAT 文件名。 |
| *Method*     | 可选          | **Variant**  | 指定附加 ExtraInfo 的方法。可以是下列 MsoExtraInfoMethod 常量之一。 |
| *HeaderInfo* | 可选          | **Variant**  | 指定 HTTP 请求的标题信息的 String。默认值为空字符串。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例加载附于第一个形状（第一张工作表中）的超链接的文档*/ Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)` |

**成员属性**

#### **Hyperlink.Address**

返回或设置一个 **String** 型，它代表目标文档的地址。

**语法**

**express.Address**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") }` |

#### **Hyperlink.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Hyperlink.EmailSubject**

返回或设置指定超链接的电子邮件主题行的文本字符串。主题行是添加到超链接地址上的。**String** 类型，可读写。

**语法**

**express.EmailSubject**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

该属性常与电子邮件超链接一起使用。

该属性的值优先于使用同一 **Hyperlink** 对象的 **Address** 属性指定的任何电子邮件主题行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例设置第一张工作表中第一个超链接的电子邮件主题行*/ Application.Worksheets.Item(1).Hyperlinks.Item(1).EmailSubject = "Quote Request"` |

#### **Hyperlink.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.Range**

返回一个 **Range** 对象，它代表附加的指定超链接的区域。

**语法**

**express.Range**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中*/ /* 示例代码*/ let rAddress = Application.Worksheets.Item("Crew").AutoFilter.Range.Addres /* 此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角*/ /* 示例代码*/ Application.Workbooks.Item(1).Activate() let hr = Application.ActiveSheet.Hyperlinks.Item(1).Range Application.ActiveWindow.ScrollRow = hr.Row Application.ActiveWindow.ScrollColumn = hr.Column` |

#### **Hyperlink.ScreenTip**

返回或设置指定超链接的屏幕提示文字。**String** 类型，可读写。

**语法**

**express.ScreenTip**

*express*   一个代表 **Hyperlink** 对象的变量。

**说明**

在将文档保存到网页中之后，如果在 Web 浏览器中查看该文档，则当鼠标移动到超链接上时，“屏幕提示”文字就会显示。某些浏览器可能不支持“屏幕提示”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例设置活动工作表上第一个超链接的屏幕提示*/ Application.ActiveSheet.Hyperlinks.Item(1).ScreenTip = "Return to the home page"` |

#### **Hyperlink.Shape**

返回一个 **Shape** 对象，它代表附加到指定超链接的形状。

**语法**

**express.Shape**

*express*   一个代表 **Hyperlink** 对象的变量。

#### **Hyperlink.SubAddress**

返回或设置文档中与指定超链接相关联的位置。**String** 类型，可读写。

**语法**

**express.SubAddress**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将一个区域位置添加到第一张形状的超链接中*/ Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.SubAddress = "A1:B10"` |

#### **Hyperlink.TextToDisplay**

返回或设置要为指定超链接显示的文本。默认值为超链接的地址。**String** 类型，可读写。

**语法**

**express.TextToDisplay**

*express*   一个代表 **Hyperlink** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为活动工作表上第一个超链接设置要显示的文本*/ Application.ActiveSheet.Hyperlinks.Item(1).TextToDisplay = "Company Home Page"` |

#### **Hyperlink.Type**

返回一个包含 **MsoHyperlinkType** 常量的 **Long** 值，它代表 HTML 框架的位置。

**语法**

**express.Type**

*express*   一个代表 **Hyperlink** 对象的变量。

适用环境：web

适用平台：windows/linux