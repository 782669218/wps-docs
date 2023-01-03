**Pane**



一个对象，它代表普通视图中的三个窗格之一，或文档窗口中任意其他视图的某个窗格。

**说明**

使用 **Panes**(*index*) 返回单个 **Pane** 对象，其中 *index* 是窗格的索引号。下表列出了普通视图中窗格的名称及其相应的索引号。

| 窗格   | 索引号 |
| ------ | ------ |
| 大纲   | 1      |
| 幻灯片 | 2      |
| 备注   | 3      |

当使用文档窗口视图而非普通视图时，使用 **Panes**(1) 可引用单个 **Pane** 对象。

使用 **Activate** 方法可激活指定窗格。

使用 **ViewType** 属性可决定活动窗格。

普通视图是唯一具有多窗格的视图。所有其他文档窗口视图都只有一个窗格，即文档窗口。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使用 ViewType 属性判断幻灯片窗格是否为活动窗格。如果幻灯片窗格是活动窗格，则使用 Activate 方法激活备注窗格。*/ function test() {     let window = Application.ActiveWindow     if(window.ActivePane.ViewType == ppViewSlide){         window.Panes.Item(3).Activate()     } }` |

**方法**

|                                                              | 名称         | 说明             |
| ------------------------------------------------------------ | ------------ | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate** | 激活指定的对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**      | 返回指定的窗格或窗口是否处于激活的状态。只读。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ViewType**    | 在带有窗格的任一视图中，返回指定窗格的视图类型。用于不带窗格的视图时，则返回父 **DocumentWindow** 对象的视图类型。只读 **PpViewType** 类型。 |

**成员方法**

#### **Pane.Activate**

激活指定的对象。

**语法**

**express.Activate()**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按顺序激活紧挨在活动窗口之后的文档窗口。*/ Application.Windows.Item(2).Activate()` |

**成员属性**

#### **Pane.Active**

返回指定的窗格或窗口是否处于激活的状态。只读。**MsoTriState** 类型。

**语法**

**express.Active**

*express*   一个代表 **Pane** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**：指定的窗格或窗口处于活动状态。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿。*/ function test() {   let mypane = Application.Presentations.Item("C:\\Users\\Administrator\\Desktop\\text.ppt").Windows.Item(1)   if(mypane.Active == false){       let oldWin = Application.ActiveWindow       mypane.Activate()   } }` |

#### **Pane.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {   let mypane = Application.ActivePresentation.Slides.Item(1).Shapes      for(let i = 1; i <= mypane.Count; i++){       if(mypane.Item(i).Type == msoLinkedOLEObject){           alert(mypane.Item(i).OLEFormat.Application.Name)       }   } }` |

#### **Pane.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **Pane.ViewType**

在带有窗格的任一视图中，返回指定窗格的视图类型。用于不带窗格的视图时，则返回父 **DocumentWindow** 对象的视图类型。只读 **PpViewType** 类型。

**语法**

**express.ViewType**

*express*   一个代表 **Pane** 对象的变量。

**说明**

| PpViewType 可以是下列 PpViewType 常量之一。 |
| ------------------------------------------- |
| **ppViewHandoutMaster**                     |
| **ppViewMasterThumbnails**                  |
| **ppViewNormal**                            |
| **ppViewNotesMaster**                       |
| **ppViewNotesPage**                         |
| **ppViewOutline**                           |
| **ppViewPrintPreview**                      |
| **ppViewSlide**                             |
| **ppViewSlideMaster**                       |
| **ppViewSlideSorter**                       |
| **ppViewThumbnails**                        |
| **ppViewTitleMaster**                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前窗格中的视图是幻灯片视图，则本示例使备注窗格成为当前窗格。备注窗格是 Panes 集合中的第三个成员。*/ function test() {   let mypane = Application.ActiveWindow   if(mypane.ActivePane.ViewType == ppViewSlide){      mypane.Panes.Item(3).Activate()   } }` |

适用环境：web

适用平台：windows/linux