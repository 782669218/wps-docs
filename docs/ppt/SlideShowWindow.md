**SlideShowWindow**



代表运行幻灯片放映的窗口。

**方法**

|                                                              | 名称         | 说明           |
| ------------------------------------------------------------ | ------------ | -------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate** | 激活指定的对象 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**       | 返回指定的窗格或窗口是否处于激活的状态。只读。MsoTriState 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**       | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsFullScreen** | 返回指定幻灯片放映窗口是否占用全屏。只读 MsoTriState 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**         | 返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Presentation** | 返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**          | 返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **View**         | 返回一个 [SlideShowView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideShowView/SlideShowView%20.htm#jsObject_SlideShowView) 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**        | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |

**成员方法**

#### **SlideShowWindow.Activate**

激活指定的对象

**语法**

**express.Activate()**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按顺序激活紧挨在活动窗口之后的文档窗口*/  Application.Windows.Item(1).Activate()` |

**成员属性**

#### **SlideShowWindow.Active**

返回指定的窗格或窗口是否处于激活的状态。只读。MsoTriState 类型。

**语法**

**express.Active**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse
msoTriStateMixed
msoTriStateToggle
msoTrue：指定的窗格或窗口处于活动状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿*/ function test(){     if(Application.Presentations.Item("test.ppt").Windows.Item(1).Active == false) {         let oldWin = Application.ActiveWindow         Application.Presentations.Item("test.ppt").Windows.Item(1).Activate()     } }` |

#### **SlideShowWindow.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/ function test() {     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")   }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/ function test(){    let shpOle = ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= shpOle.Count; i++) {         if(shpOle.Item(i).Type == msoLinkedOLEObject) {             MsgBox(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **SlideShowWindow.Height**

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第二个文档窗口的高度设置为应用程序窗口高度的一半*/ Application.Windows.Item(2).Height = Application.Windows.Item(1).Height/2` |

#### **SlideShowWindow.IsFullScreen**

返回指定幻灯片放映窗口是否占用全屏。只读 MsoTriState 类型。

**语法**

**express.IsFullScreen**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**说明**

MsoTriState 可以是下列 MsoTriState 常量之一。
msoCTrue
msoFalse
msoTriStateMixed
msoTriStateToggle
msoTrue 指定的幻灯片放映窗口占用全屏。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将全屏的幻灯片放映窗口的高度适当缩小以便可看到任务栏*/ function test(){     let SliWin = Application.ActivePresentation     if(SliWin.IsFullScreen) {         SliWin.Height = SliWin.Height - 20     } }` |

#### **SlideShowWindow.Left**

返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

*express*   一个代表 **SlideShowWindow** 对象的变量。

#### **SlideShowWindow.Presentation**

返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。

**语法**

**express.Presentation**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**说明**

如果在第一个文档窗口中当前放映的幻灯片来自嵌入的演示文稿，则 Windows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 Windows(1).Presentation 返回在其中创建第一个文档窗口的演示文稿。

如果在第一个幻灯片放映窗口中当前放映的幻灯片来自嵌入的演示文稿，则 SlideShowWindows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 SlideShowWindows(1).Presentation 返回在其中创建第一个幻灯片放映窗口的演示文稿。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使第一个窗口中的演示文稿幻灯片编号延续到第二个窗口中*/ function test(){     let firstPresSlides =  Application.Windows.Item(1).Presentation.Slides.Count      Application.Windows.Item(2).Presentation.PageSetup.FirstSlideNumber = firstPresSlides + 1 }` |

#### **SlideShowWindow.Top**

返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口*/ function test(){     let sngHeight = Application.Windows.Item(1).Height                     // available height     let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width      let dWindows1 = Application.Windows.Item(1)         dWindows1.Width = sngWidth         dWindows1.Height = sngHeight / 2         dWindows1.Left = 0      let dWindows2 = Application.Windows.Item(2)         dWindows2.Width = sngWidth         dWindows2.Height = sngHeight / 2         dWindows2.Top = sngHeight / 2         dWindows2.Left = 0 }` |

#### **SlideShowWindow.View**

返回一个 [SlideShowView](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideShowView/SlideShowView%20.htm#jsObject_SlideShowView) 对象。只读。

**语法**

**express.View**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前窗口中的视图设为幻灯片视图，然后显示第三张幻灯片*/ function test(){     Application.ActiveWindow.ViewType = ppViewSlide     Application.ActiveWindow.View.GotoSlide(3) }` |

#### **SlideShowWindow.Width**

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

*express*   一个代表 **SlideShowWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口*/ function test(){     let sngHeight = Application.Windows.Item(1).Height                     // available height     let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width      let dWindows1 = Application.Windows.Item(1)         dWindows1.Width = sngWidth         dWindows1.Height = sngHeight / 2         dWindows1.Left = 0      let dWindows2 = Windows.Item(2)         dWindows2.Width = sngWidth         dWindows2.Height = sngHeight / 2         dWindows2.Top = sngHeight / 2         dWindows2.Left = 0 }` |

适用环境：web

适用平台：windows/linux