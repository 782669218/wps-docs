**Pane**



代表窗口中的窗格。

**说明**

**Pane** 对象只对于工作表和 ET 4.0 宏表存在。**Pane** 对象是 **Panes** 集合的成员。**Panes** 集合包含在一个窗口中显示的所有窗格。

使用 **Panes**(*index* )（其中 *index* 是窗格索引号）可返回一个 **Pane** 对象。下例拆分显示工作表一的窗口，然后滚动左下角的窗格，直至第五行处于该窗格的顶端。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){         Application.Worksheets.Item(1).Activate()     Application.ActiveWindow.Split = true     Application.ActiveWindow.Panes.Item(3).ScrollRow = 5 }` |

**方法**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**              | 激活窗格。返回Boolean值                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LargeScroll**           | 按页滚动窗口内容。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToScreenPixelsX** | 返回或设置屏幕上的像素点。返回Long值                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToScreenPixelsY** | 返回或设置屏幕像素的位置。返回Long值                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScrollIntoView**        | 滚动文档窗口，使指定矩形区域中的内容显示在文档窗口或窗格的左上角或右下角（取决于 *Start* 参数值）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SmallScroll**           | 按行或按列滚动窗口内容。返回Variant值                        |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**        | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScrollColumn** | 返回或设置指定窗格或窗口最左边的列号。**Long** 型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScrollRow**    | 返回或设置指定窗格或窗口最上面显示的行号。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VisibleRange** | 返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。 |

**成员方法**

#### **Pane.Activate**

激活窗格。返回Boolean值

**语法**

**express.Activate()**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.LargeScroll**

按页滚动窗口内容。

**语法**

**express.LargeScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**             |
| --------- | ------------- | ------------ | -------------------- |
| *Down*    | 可选          | **Variant**  | 向下滚动内容的页数。 |
| *Up*      | 可选          | **Variant**  | 向上滚动内容的页数。 |
| *ToRight* | 可选          | **Variant**  | 向右滚动内容的页数。 |
| *ToLeft*  | 可选          | **Variant**  | 向左滚动内容的页数。 |

**说明**

如果同时指定了 *Down* 和 *Up*，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *Down* 为 3，*Up* 为 6，则窗口向上滚动三页。

如果同时指定了 *ToLeft* 和 *ToRight*，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *ToLeft* 为 3，*ToRight* 为 6，则窗口向右滚动三页。

以上任何参数均可为负数。

#### **Pane.PointsToScreenPixelsX**

返回或设置屏幕上的像素点。返回Long值

**语法**

**express.PointsToScreenPixelsX(Points)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**         |
| -------- | ------------- | ------------ | ---------------- |
| *Points* | 必选          | **Long**     | 屏幕像素的位置。 |

#### **Pane.PointsToScreenPixelsY**

返回或设置屏幕像素的位置。返回Long值

**语法**

**express.PointsToScreenPixelsY(Points)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Points* | 必选          | **Long**     | 起始点的位置。 |

#### **Pane.ScrollIntoView**

滚动文档窗口，使指定矩形区域中的内容显示在文档窗口或窗格的左上角或右下角（取决于 *Start* 参数值）。

**语法**

**express.ScrollIntoView(Left, Top, Width, Height, Start)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Left*   | 必选          | **Long**     | 矩形距离文档窗口或窗格左边的水平位置（以磅为单位）。         |
| *Top*    | 必选          | **Long**     | 矩形距离文档窗口或窗格上边的垂直位置（以磅为单位）。         |
| *Width*  | 必选          | **Long**     | 矩形的宽度（以磅为单位）。                                   |
| *Height* | 必选          | **Long**     | 矩形的高度（以磅为单位）。                                   |
| *Start*  | 可选          | **Variant**  | 如果为 True，则使矩形的左上角位于文档窗口或窗格的左上角。如果为 False，则使矩形的右下角位于文档窗口或窗格的右下角。默认值是 True。 |

**说明**

当矩形比文档窗口或窗格大时，*Start* 参数对调整屏幕显示很有用处。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在当前活动文档窗口中定义一个 100x200 像素的矩形，其位置为距窗口顶部 20 像素，距窗口左边缘 50 像素。然后将文档向左、向上滚动，以使矩形的左上角与窗口的左上角对齐。*/ Application.ActiveWindow.ScrollIntoView(50,20,100,200)` |

#### **Pane.SmallScroll**

按行或按列滚动窗口内容。返回Variant值

**语法**

**express.SmallScroll(Down, Up, ToRight, ToLeft)**

*express*   一个代表 **Pane** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**               |
| --------- | ------------- | ------------ | ---------------------- |
| *Down*    | 可选          | **Variant**  | 将内容向下滚动的行数。 |
| *Up*      | 可选          | **Variant**  | 将内容向上滚动的行数。 |
| *ToRight* | 可选          | **Variant**  | 将内容向右滚动的列数。 |
| *ToLeft*  | 可选          | **Variant**  | 将内容向左滚动的列数。 |

**说明**

如果同时指定了 *Down* 和 *Up*，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *Down* 为 3，*Up* 为 6，则窗口向上滚动三行。

如果同时指定了 *ToLeft* 和 *ToRight*，窗口内容的滚动量由这两个参数的差值决定。例如，如果 *ToLeft* 为 3，*ToRight* 为 6，则窗口内容向右滚动三列。

任何参数均可为负数。

**成员属性**

#### **Pane.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ if(Application.ActiveWorkbook.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") }` |

#### **Pane.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Pane** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Pane.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Pane** 对象的变量。

#### **Pane.ScrollColumn**

返回或设置指定窗格或窗口最左边的列号。**Long** 型，可读写。

**语法**

**express.ScrollColumn**

*express*   一个代表 **Pane** 对象的变量。

**说明**

如果窗口处于拆分状态，则 **Window** 对象的 **ScrollColumn** 属性表示左上方的窗格。如果某些窗格处于冻结状态，则 **Window** 对象的 **ScrollColumn** 属性将不包括冻结的区域。

#### **Pane.ScrollRow**

返回或设置指定窗格或窗口最上面显示的行号。**Long** 型，可读写。

**语法**

**express.ScrollRow**

*express*   一个代表 **Pane** 对象的变量。

**说明**

如果窗口处于拆分状态，则 **Window** 对象的 **ScrollRow** 属性表示左上方的窗格。如果某些窗格处于冻结状态，则 **Window** 对象的 **ScrollRow** 属性将不包括冻结的区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例将第十行移动到窗口顶部。*/   Application.Worksheets.Item("Sheet1").Activate()   Application.ActiveWindow.ScrollRow = 10 }` |

#### **Pane.VisibleRange**

返回一个 **Range** 对象，它代表显示在窗口或窗格中的单元格区域。如果列或行只显示了一部分，则说明它是包括在区域之内的。只读。

**语法**

**express.VisibleRange**

*express*   一个代表 **Pane** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例显示工作表 sheet1 中可见单元格的数目。*/   Application.Worksheets.Item("Sheet1").Activate()   alert("There are " + Application.Windows.Item(1).VisibleRange.Cells.Count + " cells visible") }` |

适用环境：web

适用平台：windows/linux