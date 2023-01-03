**DocumentWindow**



代表一个文档窗口。DocumentWindow对象是[DocumentWindows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/DocumentWindows/DocumentWindows%20.htm#jsObject_DocumentWindows)集合的成员。[DocumentWindows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/DocumentWindows/DocumentWindows%20.htm#jsObject_DocumentWindows)集合包含所有打开的文档窗口。

**说明**

使用 Presentation 属性可返回指定文档窗口中当前运行的演示文稿。
使用 Selection 属性可返回选定范围。
使用 SplitHorizontal 属性可返回普通视图中大纲窗格所占屏幕宽度的百分比。
使用 SplitVertical 属性可返回普通视图中幻灯片窗格所占屏幕高度的百分比。
使用 View 属性可返回指定文档窗口中的视图。

**方法**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**              | 激活指定的对象                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**                 | 关闭指定的文档窗口、演示文稿或打开的任意多边形。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FitToPage**             | 调整指定文档窗口的大小，以适应当前显示的信息。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LargeScroll**           | 滚动到指定的文档窗口中的页面。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**             | 打开一个包含指定窗口中所显示文档的新窗口。返回一个代表新窗口的 DocumentWindow 对象 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToScreenPixelsX** | 将横向度量值的单位由磅值转换为像素。可用于返回文本框或形状的横向屏幕位置。以类型 **Single** 返回转换后的度量值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToScreenPixelsY** | 将纵向度量值的单位由磅值转换为像素。可用于返回文本框或形状的纵向屏幕位置。以类型 **Single** 返回转换后的度量值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RangeFromPoint**        | 返回位于某一点（由屏幕上的位置坐标对指定）的 Shape 对象。如果在指定的坐标对上没有形状，则此方法将返回 Nothing。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScrollIntoView**        | 滚动文档窗口，使指定矩形区域内的项目显示在文档窗口或窗格中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SmallScroll**           | 按行和列滚动指定的文档窗口。                                 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**       | 返回指定的窗格或窗口是否处于激活的状态。只读MsoTriState 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePane**   | 返回一个代表文档窗口中活动窗格的Pane对象。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个Application对象，该对象表示指定对象的创建者。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**      | 返回在文档窗口标题栏中显示的文本。String类型，只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**       | 以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**         | 返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Panes**        | 返回一个 Panes 集合，该集合代表文档窗口中的 窗格 （窗格：文档窗口的一部分，以垂直或水平条为界限并由此与其他部分分隔开。） 只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Presentation** | 返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**    | 返回一个 Selection 对象，该对象代表指定的文档窗口中的选定内容。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**          | 返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **View**         | 返回一个代表指定文档窗口中视图的View对象。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ViewType**     | 返回或设置指定文档窗口中所包含的视图类型。可读/写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**        | 以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**  | 返回或设置指定窗口的状态。可读/写。PpWindowState 类型。      |

**成员方法**

#### **DocumentWindow.Activate**

激活指定的对象

**语法**

**express.Activate()**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将激活第一个文档窗口。*/ Application.Windows.Item(1).Activate()` |

#### **DocumentWindow.Close**

关闭指定的文档窗口、演示文稿或打开的任意多边形。

**语法**

**express.Close()**

*express*   一个代表 **DocumentWindow** 对象的变量。

**说明**

使用此方法，WPP关闭打开的演示文档，并且不提示用户保存所做的工作。为防止工作丢失，您应在使用 SaveClose 方法之前使用 SaveAsSave 或 CloseSaveAs 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口*/ function test(){     let dClose = Application.Windows     for ( let i = 2; i <= dClose.Count; i++ ) {         dClose.Item(i).Close()     } }   /*本示例在不保存更改的情况下关闭 Pres1.ppt*/ function test(){     let dClose1 = Application.Presentations.Item("pres1.ppt")     dClose1.Saved = true     dClose1.Close() } ` |

#### **DocumentWindow.FitToPage**

调整指定文档窗口的大小，以适应当前显示的信息。

**语法**

**express.FitToPage()**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将活动窗口中的视图设置为幻灯片视图，将显示比例设置为 25%，然后调整窗口大小以适应所显示的幻灯片。 function test() {     let dWindow = Application.ActiveWindow     dWindow.View.Zoom = 25     dWindow.FitToPage() }` |

#### **DocumentWindow.LargeScroll**

滚动到指定的文档窗口中的页面。

**语法**

**express.LargeScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**               |
| --------- | ------------- | ------------ | ---------------------- |
| *Down*    | 可选          | **Long**     | 指定要向下滚动的页数。 |
| *Up*      | 可选          | **Long**     | 指定要向上滚动的页数。 |
| *ToRight* | 可选          | **Long**     | 指定要向右滚动的页数。 |
| *ToLeft*  | 可选          | **Long**     | 指定要向左滚动的页数。 |

**说明**

如果未指定任何参数，则该方法向下滚动一页。如果同时指定了 **Down** 和 **Up**，则效果是二者的叠加。例如，如果 **Down** 为 2 而 **Up** 为 4，则该方法向上滚动两页。同样，如果同时指定了 **Right** 和 **Left**，则它们将共同起作用。

以上任何参数均可为负数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将活动窗口向下滚动三页。 Application.ActiveWindow.LargeScroll(3)` |

#### **DocumentWindow.NewWindow**

打开一个包含指定窗口中所显示文档的新窗口。返回一个代表新窗口的 DocumentWindow 对象

**语法**

**express.NewWindow()**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个与当前窗口有相同内容的新窗口（同时激活新窗口）并返回到原窗口*/ function test(){     let oldW = Application.ActiveWindow     let newW = oldW.NewWindow()     oldW.Activate() }` |

#### **DocumentWindow.PointsToScreenPixelsX**

将横向度量值的单位由磅值转换为像素。可用于返回文本框或形状的横向屏幕位置。以类型 **Single** 返回转换后的度量值。

**语法**

**express.PointsToScreenPixelsX(Points)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                         |
| -------- | ------------- | ------------ | ------------------------------------------------ |
| *Points* | 必选          | **Single**   | 要转换为以像素为单位的横向度量值（以磅为单位）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将选中的文本框架边界框的宽度和高度从磅值转换到像素。并将值返回到 myXparm 和 myYparm。 function test() {     let myXparm = Application.ActiveWindow.PointsToScreenPixelsX(Application.ActiveWindow.Selection.TextRange.BoundWidth)     alert(myXparm)     let myYparm = Application.ActiveWindow.PointsToScreenPixelsY(Application.ActiveWindow.Selection.TextRange.BoundHeight)     alert(myYparm) }` |

#### **DocumentWindow.PointsToScreenPixelsY**

将纵向度量值的单位由磅值转换为像素。可用于返回文本框或形状的纵向屏幕位置。以类型 **Single** 返回转换后的度量值。

**语法**

**express.PointsToScreenPixelsY(Points)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                         |
| -------- | ------------- | ------------ | ------------------------------------------------ |
| *Points* | 必选          | **Single**   | 要转换为以像素为单位的纵向度量值（以磅为单位）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将选取的文本框架边界框的宽度和高度从磅值转换到像素。并将值返回到 myXparm 和 myYparm。 function test() {     let myXparm = Application.ActiveWindow.PointsToScreenPixelsX(Application.ActiveWindow.Selection.TextRange.BoundWidth)     alert(myXparm)     let myYparm = Application.ActiveWindow.PointsToScreenPixelsY(Application.ActiveWindow.Selection.TextRange.BoundHeight)     alert(myYparm) }` |

#### **DocumentWindow.RangeFromPoint**

返回位于某一点（由屏幕上的位置坐标对指定）的 Shape 对象。如果在指定的坐标对上没有形状，则此方法将返回 Nothing。

**语法**

**express.RangeFromPoint(x, y)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                       |
| -------- | ------------- | ------------ | ---------------------------------------------- |
| *x*      | 必选          | **Long**     | 从屏幕左边缘到该点的水平距离（以像素为单位）。 |
| *y*      | 必选          | **Long**     | 从屏幕左边缘到该点的水平距离（以像素为单位）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使用坐标 (288, 100) 向第一张幻灯片中添加一个新的五角星。然后，再将坐标由单位磅转换为像素，并使用 RangeFromPoint 方法返回一个对该新对象的引用，最后更改此五角星的填充色。 function test() {     Application.ActivePresentation.Slides.Item(1).Shapes.AddShape(msoShape5pointStar, 288, 100, 100, 72).Select()     let myPointX = Application.ActiveWindow.PointsToScreenPixelsX(288)     let myPointY = Application.ActiveWindow.PointsToScreenPixelsY(100)     let myShape = Application.ActiveWindow.RangeFromPoint(myPointX, myPointY)     myShape.Fill.ForeColor.RGB = (80, 160, 130) }` |

#### **DocumentWindow.ScrollIntoView**

滚动文档窗口，使指定矩形区域内的项目显示在文档窗口或窗格中。

**语法**

**express.ScrollIntoView(Left, Top, Width, Height, Start)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                       |
| -------- | ------------- | --------------- | ---------------------------------------------- |
| *Left*   | 必选          | **Long**        | 从文档窗口的左边缘到该矩形的水平距离（点数）。 |
| *Top*    | 必选          | **Long**        | 从文档窗口顶部到该矩形的垂直距离（点数）。     |
| *Width*  | 必选          | **Long**        | 矩形的宽度（点数）。                           |
| *Height* | 必选          | **Long**        | 矩形的高度（点数）。                           |
| *Start*  | 可选          | **MsoTriState** | 确定矩形相对于文档窗口的起始位置。             |

**说明**

如果矩形边界超过了文档窗口的大小，则 **Start** 参数可指定矩形的哪一个顶点将显示在屏幕上或具有初始焦点。该方法不能在大纲或幻灯片浏览视图中使用。

| MsoTriState 可以是下列 MsoTriState 常量之一。              |
| ---------------------------------------------------------- |
| **msoCTrue**                                               |
| **msoFalse** 矩形的右下角将显示于文档窗口的右下角。        |
| **msoTriStateMixed**                                       |
| **msoTriStateToggle**                                      |
| **msoTrue** 默认值。矩形的左上角将显示在文档窗口的左上角。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将查看一个 100x200（单位：磅）的区域，该区域距幻灯片左边缘 50 磅，距幻灯片上边缘 20 磅。该矩形的左上角位于当前文档窗口的左上角。 Application.ActiveWindow.ScrollIntoView(50, 20, 100, 200)` |

#### **DocumentWindow.SmallScroll**

按行和列滚动指定的文档窗口。

**语法**

**express.SmallScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **DocumentWindow** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**             |
| --------- | ------------- | ------------ | -------------------- |
| *Down*    | 可选          | **Long**     | 指定向下滚动的行数。 |
| *Up*      | 可选          | **Long**     | 指定向上滚动的行数。 |
| *ToRight* | 可选          | **Long**     | 指定向右滚动的列数。 |
| *ToLeft*  | 可选          | **Long**     | 指定向左滚动的列数。 |

**说明**

如果未指定任何参数，则该方法向下滚动一行。如果同时指定了 **Down** 和 **Up**，则效果是二者的叠加。例如，如果 **Down** 为 2，而 **Up** 为 4，该方法将向上滚动两行。同样，如果同时指定了 **Right** 和 **Left**，则它们将共同起作用。

以上任何参数均可为负数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将当前窗口向下滚动三行。 Application.ActiveWindow.SmallScroll(3)` |

**成员属性**

#### **DocumentWindow.Active**

返回指定的窗格或窗口是否处于激活的状态。只读MsoTriState 类型。

**语法**

**express.Active**

*express*   一个代表 **DocumentWindow** 对象的变量。

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
| `/*以下示例检查演示文稿文件“test.ppt”是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量oldWin中，并激活“test.ppt”演示文稿*/ function test(){     let dActive = Application.Presentations.Item("test.ppt").Windows.Item(1)     if(!dActive.Active){         let oldWin = Application.ActiveWindow         dActive.Activate()     } }` |

#### **DocumentWindow.ActivePane**

返回一个代表文档窗口中活动窗格的Pane对象。只读。

**语法**

**express.ActivePane**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动窗格为幻灯片窗格，则本示例将设置备注窗格为活动窗格。备注窗格是Panes集合中的第三个成员。*/ function test(){     if(Application.ActiveWindow.ActivePane.ViewType == ppViewSlide){         Application.ActiveWindow.Panes.Item(3).Activate()     } }` |

#### **DocumentWindow.Application**

返回一个Application对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个Presentation对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行WPP的文件夹中。*/ function test(){     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1,1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

#### **DocumentWindow.Caption**

返回在文档窗口标题栏中显示的文本。String类型，只读。

**语法**

**express.Caption**

*express*   一个代表 **DocumentWindow** 对象的变量。

#### **DocumentWindow.Height**

以磅为单位返回或设置指定对象的高度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Height**

*express*   一个代表 **DocumentWindow** 对象的变量。

**说明**

一个 HeightShape 对象的 ShapeHeight 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第二文档窗口的高度设置为应用程序窗口高度的一半。*/ Application.Windows.Item(2).Height = Application.Windows.Item(1).Height/2   ` |

#### **DocumentWindow.Left**

返回或设置一个 Single 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

*express*   一个代表 **DocumentWindow** 对象的变量。

#### **DocumentWindow.Panes**

返回一个 Panes 集合，该集合代表文档窗口中的 窗格 （窗格：文档窗口的一部分，以垂直或水平条为界限并由此与其他部分分隔开。） 只读。

**语法**

**express.Panes**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检测活动窗口中窗格的数目。如果其值为1，则表明是非普通视图的任意其他视图，然后本示例将激活普通视图。*/ Application.ActiveWindow.Panes.Count ` |

#### **DocumentWindow.Presentation**

返回一个 Presentation 对象，该对象代表创建指定文档窗口或幻灯片放映窗口的演示文稿。只读。

**语法**

**express.Presentation**

*express*   一个代表 **DocumentWindow** 对象的变量。

**说明**

如果在第一个文档窗口中当前放映的幻灯片来自嵌入的演示文稿，则 Windows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 Windows(1).Presentation 返回在其中创建第一个文档窗口的演示文稿。
如果在第一个幻灯片放映窗口中当前放映的幻灯片来自嵌入的演示文稿，则 SlideShowWindows(1).View.Slide.Parent 返回该嵌入演示文稿，并且 SlideShowWindows(1).Presentation 返回在其中创建第一个幻灯片放映窗口的演示文稿。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使第一个窗口的演示文稿幻灯片编号延续到第二个窗口中。*/ function test(){     let firstPresSlides = Application.Windows.Item(1).Presentation.Slides.Count     Application.Windows.Item(2).Presentation.PageSetup.FirstSlideNumber  += 1 }` |

#### **DocumentWindow.Selection**

返回一个 Selection 对象，该对象代表指定的文档窗口中的选定内容。只读。

**语法**

**express.Selection**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果在活动窗口中选择了文本，本示例则会将该文本设为斜体。*/ function test(){     let dText = Application.ActiveWindow.Selection     if(dText.Type == ppSelectionText){         dText.TextRange.Font.Italic = true     } }` |

#### **DocumentWindow.Top**

返回或设置一个 Single 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所用水平可用空间。*/ /*要使以下示例执行，只能有两个打开的文档窗口。*/ function test(){         let sngHeight = Application.Windows.Item(1).Height     let sngWidth = Application.Windows.Item(1).Width + Windows.Item(2).Width     let dWindows1 = Application.Windows.Item(1)     dWindows1.Width = sngWidth     dWindows1.Height = sngHeight/2     dWindows1.Left = 0      let dWindows2 = Application.Windows.Item(2)     dWindows2.Width = sngWidth     dWindows2.Height = sngHeight/2     dWindows2.Top = sngHeight/2     dWindows2.Left = 0 }` |

#### **DocumentWindow.View**

返回一个代表指定文档窗口中视图的View对象。只读。

**语法**

**express.View**

*express*   一个代表 **DocumentWindow** 对象的变量。

#### **DocumentWindow.ViewType**

返回或设置指定文档窗口中所包含的视图类型。可读/写。

**语法**

**express.ViewType**

*express*   一个代表 **DocumentWindow** 对象的变量。

**说明**

PpViewType 可以是下列 PpViewType 常量之一。
ppViewHandoutMaster
ppViewMasterThumbnails
ppViewNormal
ppViewNotesMaster
ppViewNotesPage
ppViewOutline
ppViewPrintPreview
ppViewSlide
ppViewSlideMaster
ppViewSlideSorter
ppViewThumbnails
ppViewTitleMaster

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前窗口以普通视图显示，本示例将该窗口中的视图更改为幻灯片浏览视图。*/ function test(){     let dView = Application.ActiveWindow     if ( dView.ViewType == ppViewNormal ) { 	dView.ViewType = ppViewSlideSorter     } }` |

#### **DocumentWindow.Width**

以磅为单位返回或设置指定对象的宽度。用于 Master 对象时只读，Single 类型；用于所有其他对象时可读/写，Single 类型。

**语法**

**express.Width**

*express*   一个代表 **DocumentWindow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例水平排列第一个和第二个窗口。即每个窗口占据应用程序窗口用户区域的一半垂直可用空间和所有水平可用空间。要使以下示例执行，只能有两个打开的文档窗口。*/ function test(){     let ah = Application.Windows.Item(1).Height                      // available height     let aw = Application.Windows.Item(1).Width + Windows.Item(2).Width    // available width      let dWindows1 = Application.Windows.Item(1)     dWindows1.Width = aw     dWindows1.Height = ah / 2     dWindows1.Left = 0      let dWindows2 = Application.Windows.Item(2)     dWindows2.Width = aw     dWindows2.Height = ah / 2     dWindows2.Top = ah / 2     dWindows2.Left = 0 }  /*以下示例将指定表中第一列的宽度设置为 80 磅。*/ Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).Width = 80     ` |

#### **DocumentWindow.WindowState**

返回或设置指定窗口的状态。可读/写。PpWindowState 类型。

**语法**

**express.WindowState**

*express*   一个代表 **DocumentWindow** 对象的变量。

**说明**

PpWindowState 可以是下列 PpWindowState 常量之一。
ppWindowMaximized
ppWindowMinimized
ppWindowNormal
当窗口状态为 ppWindowNormal 时，该窗口既未最大化也未最小化。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前窗口最大化。*/ Application.ActiveWindow.WindowState = ppWindowMaximized` |

适用环境：web

适用平台：windows/linux