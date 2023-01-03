#### **Page**



代表文档中的页面。使用 **Page** 对象及相关方法和属性可通过编程方式定义文档的页面版式。

**说明**

使用 **Item** 方法可访问文档中的特定页面。以下示例访问活动文档的第一页。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let objPage = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1)` |

要访问页码，请使用 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)或 [**Selection** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection)对象的 **Information** 属性，或者 **Break** 对象（该对象属于指定 **Page** 对象的 **Breaks** 集合）的 **PageIndex** 属性。

[**Page** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#jsObject_Page)对象的 [**Top** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Top)和 [**Left** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Left)属性总是返回 0（零），代表页面的左上角。[**Height** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Height)和 [**Width** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Width)属性返回在“页面设置”对话框中指定的或通过 [**PageSetup** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/PageSetup/PageSetup%20.htm#jsObject_PageSetup)对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面，[**Height** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Height)属性返回 792，[**Width** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Page/Page%20.htm#Page.Width)属性返回 612。所有这四个属性都是只读的。

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Breaks**          | 返回一个 **Breaks** 集合，该集合代表页面上的分隔符。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnhMetaFileBits** | 返回一个 **Variant** 类型的值，该值代表文本页显示方式的图片表示形式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**          | 返回一个 **Long** 类型的值，该值代表页面的高度（以像素为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**            | 返回一个 **Long** 类型的值，该值代表页面的左边缘。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回一个 **Object** 类型值，该值代表指定 **Page** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rectangles**      | 返回一个 **Rectangles** 集合，该集合代表文档页面上的文字或图形部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**             | 返回一个 **Long** 类型的值，该值代表页面的上边缘。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**           | 返回一个 **Long** 类型的值，该值代表在**“页面设置”**对话框中定义的纸张宽度，以磅为单位。只读 **Long** 类型。 |

**成员属性**

#### **Page.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Page** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Page.Breaks**

返回一个 **Breaks** 集合，该集合代表页面上的分隔符。

**语法**

**express.Breaks**

*express*   一个代表 **Page** 对象的变量。

**说明**

**Breaks** 集合包含分页符、分栏符和分节符。使用 **Breaks** 集合以及相关对象和属性可通过编程方式定义文档的页面版式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中第一页上的分隔符。*/ let objBreaks = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Breaks` |

#### **Page.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Page** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Page.EnhMetaFileBits**

返回一个 **Variant** 类型的值，该值代表文本页显示方式的图片表示形式。只读。

**语法**

**express.EnhMetaFileBits**

*express*   一个代表 **Page** 对象的变量。

**说明**

**EnhMetaFileBits** 属性返回一个字节数组，该数组可在 Visual Basic 或 Microsoft C++ 开发环境中用于 Microsoft Windows 32 应用程序编程接口。

#### **Page.Height**

返回一个 **Long** 类型的值，该值代表页面的高度（以像素为单位）。

**语法**

**express.Height**

*express*   一个代表 **Page** 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），以指示页面的左上角。**Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 [PageSetup ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/PageSetup/PageSetup%20.htm#jsObject_PageSetup)对象指定的纸张大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面，**Height** 属性返回 792，**Width** 属性返回 612。所有这四个属性都是只读的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Page 对象的 Top 和 Left 属性总是返回 0（零），以指示页面的左上角。 Height 和 Width 属性返回在“页面设置”对话框中指定的或通过 PageSetup 对象指定的纸张大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。 例如，对于纵向模式下的 8-1/2 × 11 英寸页面，Height 属性返回 792，Width 属性返回 612。所有这四个属性都是只读的。` |

#### **Page.Left**

返回一个 **Long** 类型的值，该值代表页面的左边缘。只读。

**语法**

**express.Left**

*express*   一个代表 **Page** 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。**Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面，**Height** 属性返回 792，**Width** 属性返回 612。所有这四个属性都是只读的。

#### **Page.Parent**

返回一个 **Object** 类型值，该值代表指定 **Page** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Page** 对象的变量。

#### **Page.Rectangles**

返回一个 **Rectangles** 集合，该集合代表文档页面上的文字或图形部分。

**语法**

**express.Rectangles**

*express*   一个代表 **Page** 对象的变量。

**说明**

使用 **Rectangles** 集合及相关对象和属性可通过编程方式定义文档的页面版式。矩形对应于文档页面上的文字或图形部分。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中第一页的 Rectangles 集合。*/ let objRectangles = ActiveDocument.ActiveWindow.Panes.Item(1).Pages.Item(1).Rectangles` |

#### **Page.Top**

返回一个 **Long** 类型的值，该值代表页面的上边缘。只读。

**语法**

**express.Top**

*express*   一个代表 **Page** 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。**Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面，**Height** 属性返回 792，**Width** 属性返回 612。所有这四个属性都是只读的。

#### **Page.Width**

返回一个 **Long** 类型的值，该值代表在**“页面设置”**对话框中定义的纸张宽度，以磅为单位。只读 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **Page** 对象的变量。

**说明**

**Page** 对象的 **Top** 和 **Left** 属性总是返回 0（零），代表页面的左上角。**Height** 和 **Width** 属性返回在“页面设置”对话框中指定的或通过 **PageSetup** 对象指定的页面大小的高度和宽度，以磅为单位（72 磅 = 1 英寸）。例如，对于纵向模式下的 8-1/2 × 11 英寸页面，**Height** 属性返回 792，**Width** 属性返回 612。所有这四个属性都是只读的。

适用环境：web

适用平台：windows/linux