**Selection**



代表指定文档窗口中的选定范围。每次在幻灯片视图中更改幻灯片时，Selection 对象将被删除（Type 属性将返回 ppSelectionNone）。

**方法**

|                                                              | 名称         | 说明                             |
| ------------------------------------------------------------ | ------------ | -------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**     | 将指定对象复制到剪贴板。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**      | 删除指定对象并将其放到剪贴板中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**   | 删除指定的对象。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unselect** | 取消当前选择。                   |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildShapeRange** | 返回一个代表选定范围的子形状的 ShapeRange 对象。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeRange**      | 返回一个 ShapeRange 对象，该对象代表指定幻灯片上所有选定的幻灯片对象。其范围可包括幻灯片上的图形、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号占位符以及日期和时间对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlidRange**       | 返回一个 SlideRange 对象，该对象代表选定的幻灯片范围。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextRange**       | 返回一个 TextRange 对象，该对象代表选定文本（[Selection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection) 对象）或指定文本框（[TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextFrame/TextFrame%20.htm#jsObject_TextFrame) 对象）中的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回一个 PpSelectionType 常量，该常量代表选定范围中对象的类型。只读。 |

**成员方法**

#### **Selection.Copy**

将指定对象复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用 **Paste** 方法粘贴剪贴板中的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一个窗口中拷贝选定幻灯片到剪贴板然后将其粘贴到第二窗口的视图中。如果剪贴板的内容不能粘贴到第二个窗口的视图中（例如，如果试图将一个形状粘贴到幻灯片浏览视图中），则本示例失败*/ function test(){     Application.Windows.Item(1).Selection.Copy()     Application.Windows.Item(2).View.Paste() }  /*本示例将当前演示文稿的第一张幻灯片复制到剪贴板*/ Application.ActivePresentation.Slides.Item(1).Copy()   /*本示例将当前演示文稿中第一张幻灯片的第一个形状中的文本复制到剪贴板*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Copy() ` |

#### **Selection.Cut**

删除指定对象并将其放到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象并将它的一个副本放到剪贴板中*/  Application.Windows.Item(1).Selection.Cut()  /*本示例从当前演示文稿的第一张幻灯片中删除第一个和第二个形状，并将它们的副本放到剪贴板，然后将该副本粘贴到第二张幻灯片*/ function test(){     let active = Application.ActivePresentation         active.Slides.Item(1).Shapes.Range([1, 2]).Cut()         active.Slides.Item(2).Shapes.Paste() }   /*本示例从当前演示文稿中删除第一张幻灯片并将它的一个副本放到剪贴板中*/ Application.ActivePresentation.Slides.Item(1).Cut()   /*本示例删除当前演示文稿中第一张幻灯片的第一个形状中的文本，并将它的一个副本放到剪贴板中*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Cut()` |

#### **Selection.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除第一个窗口中选中的对象 Application.Windows.Item(1).Selection.Delete()` |

#### **Selection.Unselect**

取消当前选择。

**语法**

**express.Unselect()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消对第一个窗口中当前选定内容的选定*/  Application.Windows.Item(1).Selection.Unselect()` |

**成员属性**

#### **Selection.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/ function test(){     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/ function test(){     let shpole =Application.ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= shpole.Count; i++){         if(shpole.Item(i).Type == msoLinkedOLEObject){             MsgBox(shpole.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **Selection.ChildShapeRange**

返回一个代表选定范围的子形状的 ShapeRange 对象。

**语法**

**express.ChildShapeRange**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ShapeRange**

返回一个 ShapeRange 对象，该对象代表指定幻灯片上所有选定的幻灯片对象。其范围可包括幻灯片上的图形、形状、OLE 对象、图片、文本对象、标题、页眉、页脚、幻灯片编号占位符以及日期和时间对象。只读。

**语法**

**express.ShapeRange**

*express*   一个代表 **Selection** 对象的变量。

**说明**

当演示文稿为普通视图、幻灯片视图或任一母版视图时，可以从选定范围中返回一个形状范围。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置第一个窗口中所有选定形状的前景填充颜色*/ Application.Windows.Item(1).Selection.ShapeRange.Fill.ForeColor.RGB = (255,0,255)` |

#### **Selection.SlidRange**

返回一个 SlideRange 对象，该对象代表选定的幻灯片范围。只读。

**语法**

**express.SlidRange**

*express*   一个代表 **Selection** 对象的变量。

**说明**

可以在幻灯片视图、幻灯片浏览视图、普通视图、备注页视图或大纲视图中构造幻灯片范围。在幻灯片视图中，SlideRange 返回当前放映的幻灯片。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置第一个窗口中所有选定幻灯片的数量*/ Application.Windows.Item(1).Selection.SlideRange.Count` |

#### **Selection.TextRange**

返回一个 TextRange 对象，该对象代表选定文本（[Selection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection) 对象）或指定文本框（[TextFrame](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/TextFrame/TextFrame%20.htm#jsObject_TextFrame) 对象）中的文本。

**语法**

**express.TextRange**

*express*   一个代表 **Selection** 对象的变量。

**说明**

当演示文稿在幻灯片视图、普通视图、大纲视图、备注页视图或任何母版视图时，可从选定的内容中创建文本范围。

#### **Selection.Type**

返回一个 PpSelectionType 常量，该常量代表选定范围中对象的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **Selection** 对象的变量。

**说明**

PpSelectionType 可以是下列 PpSelectionType 常量之一。
ppSelectionNone
ppSelectionShapes
ppSelectionSlides
ppSelectionText

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将显示选中对象的类型*/ Application.Windows.Item(1).Selection.Type` |

适用环境：web

适用平台：windows/linux