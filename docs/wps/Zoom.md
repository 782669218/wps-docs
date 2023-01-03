#### **Zoom**



包含窗口或窗格的显示比例选项（例如，显示比例）。**Zoom** 对象是 **Zooms** 集合的一个成员。

**说明**

使用 **View** 对象的 **Zoom** 属性可返回一个 **Zoom** 对象。以下示例将活动窗口的显示比例设置为 110%。

| 示例代码复制                                             |
| -------------------------------------------------------- |
| `ActiveDocument.ActiveWindow.View.Zoom.Percentage = 110` |

使用 **Zooms**(*Index*) 可返回一个 **Zoom** 对象，其中 *Index* 标识视图类型。由 index 指定的视图类型可以是下列 [WdViewType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)常量之一：[wdMasterView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)、[wdNormalView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)、**wdOutlineView**、[wdPrintPreview](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)、[**wdPrintView** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)或 **wdWebView**。以下示例设置活动窗口的缩放比例，以便看到整个页面。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `ActiveDocument.ActiveWindow.ActivePane.Zooms.Item(wdPrintView).PageFit = wdPageFitFullPage` |

**Add** 方法不可用于 **Zooms** 集合。对于每个不同的视图类型（如大纲视图、普通视图或页面视图），**Zooms** 集合都只包含一个 **Zoom** 对象。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageColumns** | 返回或设置需要在页面视图或打印预览中在屏幕上并排显示的页数。Long 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageFit**     | 返回或设置窗口的显示比例，以便能查看整页或整个页宽的范围。**WdPageFit** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageRows**    | 返回或设置在页面视图或打印预览中在屏幕上垂直显示的页数。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Zoom** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Percentage**  | 用百分比的形式返回或设置窗口的显示比例。**Long** 类型，可读写。 |

**成员属性**

#### **Zoom.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Zoom** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Zoom.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Zoom** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Zoom.PageColumns**

返回或设置需要在页面视图或打印预览中在屏幕上并排显示的页数。Long 类型，可读写。

**语法**

**express.PageColumns**

*express*   一个代表 **Zoom** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到页面视图，并显示并列竖排的两页。*/ function test() {     let view = Application.ActiveDocument.ActiveWindow.View     view.Type = wdPrintView     view.Zoom.PageColumns = 2     view.Zoom.PageRows = 1 }  /*本示例将 Hello.doc 的文档窗口切换到页面视图，并显示整页。*/ function test() {     let view = Application.Windows.Item("Hello.doc").View     view.Type = wdPrintView     let z = view.Zoom     z.PageColumns = 1     z.PageRows = 1     z.PageFit = wdPageFitFullPage } ` |

#### **Zoom.PageFit**

返回或设置窗口的显示比例，以便能查看整页或整个页宽的范围。**WdPageFit** 类型，可读写。

**语法**

**express.PageFit**

*express*   一个代表 **Zoom** 对象的变量。

**说明**

如果文档不在页面视图内，则 **wdPageFitFullPage** 常量无效。

如果将 **PageFit** 属性设置为 **wdPageFitBestFit**，则在每次改变文档窗口的大小时，都将自动重新计算显示比例。如果该属性设置为 **wdPageFitNone**，则在改变文档窗口的大小时，不会重新计算显示比例。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改 Letter.doc 的窗口的显示比例，以便可查看整个文本宽度的范围。*/ let view = Application.Windows.Item("Letter.doc").View view.Type = wdNormalView view.Zoom.PageFit = wdPageFitBestFit  /*本示例将活动窗口切换到页面视图并更改显示比例，以便可查看整页。*/ let view = Application.ActiveDocument.ActiveWindow.View view.Type = wdPrintView view.Zoom.PageFit = wdPageFitFullPage` |

#### **Zoom.PageRows**

返回或设置在页面视图或打印预览中在屏幕上垂直显示的页数。**Long** 类型，可读写。

**语法**

**express.PageRows**

*express*   一个代表 **Zoom** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到打印预览并垂直显示两页。*/ PrintPreview = true let zoom = Application.ActiveDocument.ActiveWindow.View.Zoom zoom.PageColumns = 1 zoom.PageRows = 2` |

#### **Zoom.Parent**

返回一个 **Object** 类型值，该值代表指定 **Zoom** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Zoom** 对象的变量。

#### **Zoom.Percentage**

用百分比的形式返回或设置窗口的显示比例。**Long** 类型，可读写。

**语法**

**express.Percentage**

*express*   一个代表 **Zoom** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口切换到普通视图并将显示比例设置为 80%。*/ let view = Application.ActiveDocument.ActiveWindow.View view.Type = wdNormalView view.Zoom.Percentage = 80  /*本示例将活动窗口的显示比例增加 10%。*/ let myZoom = Application.ActiveDocument.ActiveWindow.View.Zoom myZoom.Percentage = myZoom.Percentage + 10` |

适用环境：web

适用平台：windows/linux