**OLEObjects**



指定工作表中所有 **OLEObject** 对象的集合。

**说明**

每一个 **OLEObject** 对象都代表一个 ActiveX 控件或一个链接或嵌入的 OLE 对象。

工作表上的 ActiveX 控件有两个名称：一个是包含该控件的形状的名称，查看工作表时可在**“名称”**框中看到它；另一个是控件的代码名称，可以在**“属性”**窗口中**“(名称)”**右边的单元格中看到它。在您首次向工作表添加控件时，形状名称和代码名称是一致的。但是，如果您更改这两个名称中的任意一个，另一个不会随之自动更改。

使用 **OLEObjects** 方法可返回 **OLEObjects** 集合。下例隐藏工作表一上的所有 OLE 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).OLEObjects().Visible = false` |

使用 [**Add** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/OLEObjects/OLEObjects%20.htm#OLEObjects.Add)方法可创建一个新 OLE 对象并将它添加到 **OLEObjects** 集合中草药。下例创建一个新 OLE 对象（代表位图文件 Arcade.bmp）并将它添加到工作表一。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).OLEObjects().Add ("arcade.gif")` |

下例新建一个 ActiveX 控件（一个列表框），并将其添加到第一张工作表中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).OLEObjects().Add ("Forms.ListBox.1")` |

在控件事件过程名称中使用的是控件的代码名称，但是，当您从工作表的 **Shapes** 或 **OLEObjects** 集合中返回控件时，必须使用形状名称而不是代码名称，以便按名称引用控件。例如，假定您给工作表添加一个复选框，而默认的形状名称和代码名称都是 CheckBox1。如果通过在**“属性”**窗口的**“(名称)”**旁边键入“chkFinished”而更改了控件代码名称，则在事件过程名称中必须使用 chkFinished，但是您仍然需要使用 CheckBox1 从 **Shapes** 或 **OLEObject** 集合中返回控件，如下例所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function chkFinished_Click(){     Application.ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 }` |

**方法**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**          | 向工作表中添加新的 OLE 对象。返回一个 **OLEObject** 对象，它代表新的 OLE 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BringToFront** | 将对象放到 z-次序前面。返回Variant值                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**         | 将对象复制到剪贴板。返回Variant值                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyPicture**  | 将所选对象作为图片复制到剪贴板。返回**Variant值**。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**          | 将对象剪切到剪贴板。返回Variant值                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。返回Variant值                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate**    | 复制对象，并返回对新复制对象的引用。返回Object值             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**         | 从集合中返回一个对象，它代表集合包含的某个对象。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择对象。返回Variant值                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendToBack**   | 将对象放到 z-次序的后面。返回Variant值                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoLoad**    | 如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**      | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**     | 如果启用对象，则为 **True**。**Boolean** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**      | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Interior**    | 返回一个 **Interior** 对象，它代表指定对象的内部。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**        | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**      | 返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Placement**   | 返回或设置一个包含 **XlPlacement**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPlacement%20%E6%9E%9A%E4%B8%BE.html)常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintObject** | 如果打印文档时也打印指定对象，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**      | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShapeRange**  | 返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceName**  | 返回或设置一个 **String** 值，它代表指定对象的链接源名称。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**         | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**       | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ZOrder**      | 返回指定对象的 z-次序位置。**Long** 型，只读。               |

**成员方法**

#### **OLEObjects.Add**

向工作表中添加新的 OLE 对象。返回一个 **OLEObject** 对象，它代表新的 OLE 对象。

**语法**

**express.Add(ClassType, Height, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Width)**

*express*   一个代表 **OLEObjects** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ClassType*     | 可选          | **Variant**  | （必须指定 ClassType 或 FileName）。一个字符串，包含要创建的对象的程序标识符。如果指定了 ClassType 参数，则忽略 FileName 和 Link。 |
| *Height*        | 可选          | **Variant**  | （必须指定 ClassType 或 FileName）。一个字符串，指定用于创建 OLE 对象的文件。 |
| *Link*          | 可选          | **Variant**  | 如果为 True，则让基于 FileName 的新 OLE 对象链接到该文件。如果该对象未链接到文件，则该对象被创建为文件副本。默认值是 False。 |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果为 True，则以图标或正常图片方式显示新的 OLE 对象。如果该参数设置为 True，则可以使用 IconFileName 和 IconIndex 来指定图标。 |
| *IconFileName*  | 可选          | **Variant**  | 一个字符串，指定要显示的图标所在的文件。仅当 DisplayAsIcon 为 True 时，才使用该参数。如果不指定该参数，或文件中不包含图标，则使用 OLE 类的默认图标。 |
| *IconIndex*     | 可选          | **Variant**  | 图标文件中包含的图标数目。仅当 DisplayAsIcon 参数为 True 并且 IconFileName 参数引用包含图标的有效文件时，才使用该参数。如果由 IconFileName 参数指定的文件中不存在具有指定索引号的图标，则使用该文件中的第一个图标。 |
| *IconLabel*     | 可选          | **Variant**  | 一个字符串，指定在图标下方显示一个标签。仅当 DisplayAsIcon 为 True 时，才使用该参数。如果省略该参数，或者该参数为空字符串 ("")，则不显示任何标题。 |
| *Left*          | 可选          | **Variant**  | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。 |
| *Width*         | 可选          | **Variant**  | 以磅为单位给出新对象的初始大小。                             |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `funtion test(){   /*在工作表 Sheet1 中新建一个 WPS OLE 对象。*/   Application.ActiveWorkbook.Worksheets.Item("Sheet1").OLEObjects().Add("WPS.Document")   /*为第一张工作表添加一个命令按钮。*/   Application.Worksheets.Item(1).OLEObjects().Add("Forms.CommandButton.1", false, false, 40, 40, 150, 10)    }` |

#### **OLEObjects.BringToFront**

将对象放到 z-次序前面。返回Variant值

**语法**

**express.BringToFront()**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Copy**

将对象复制到剪贴板。返回Variant值

**语法**

**express.Copy()**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.CopyPicture**

将所选对象作为图片复制到剪贴板。返回**Variant值**。

**语法**

**express.CopyPicture(Appearance, Format)**

*express*   一个代表 **OLEObjects** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**            | **说明**                                |
| ------------ | ------------- | ----------------------- | --------------------------------------- |
| *Appearance* | 可选          | **XlPictureAppearance** | 指定图片的复制方式。默认值为 xlScreen。 |
| *Format*     | 可选          | **XlCopyPictureFormat** | 图片的格式。默认值为 xlPicture。        |

#### **OLEObjects.Cut**

将对象剪切到剪贴板。返回Variant值

**语法**

**express.Cut()**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Delete**

删除对象。返回Variant值

**语法**

**express.Delete()**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Duplicate**

复制对象，并返回对新复制对象的引用。返回Object值

**语法**

**express.Duplicate()**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Item**

从集合中返回一个对象，它代表集合包含的某个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **OLEObjects** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 可选          | **Variant**  | 对象的名称或索引号。 |

**说明**

对象的文本名称就是 **Name** 和 **Value** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将工作表 Sheet1 中的第一个 OLE 对象删除。*/ Application.Worksheets.Item("sheet1").OLEObjects().Item(1).Delete()` |

#### **OLEObjects.Select**

选择对象。返回Variant值

**语法**

**express.Select(Replace)**

*express*   一个代表 **OLEObjects** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

#### **OLEObjects.SendToBack**

将对象放到 z-次序的后面。返回Variant值

**语法**

**express.SendToBack()**

*express*   一个代表 **OLEObjects** 对象的变量。

**成员属性**

#### **OLEObjects.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **OLEObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例显示一条有关创建 myObject 的应用程序的消息。*/   let myObject = Application.ActiveWorkbook   if (myObject.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **OLEObjects.AutoLoad**

如果打开包含指定 OLE 对象的工作簿时自动载入该 OLE 对象，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoLoad**

*express*   一个代表 **OLEObjects** 对象的变量。

**说明**

ActiveX 忽略此属性。打开一个工作簿时总会载入 ActiveX 控件。

对于大多数 OLE 对象类型，此属性不能设为 **True**。对于新 OLE 对象，默认情况下其 **AutoLoad** 属性设为 **False**；当 ET 载入工作簿时，设为“False”可节省时间和内存。自动载入 OLE 对象的好处在于，对于代表易变动的数据的对象，可立即重建到数据源的链接，而且，如果需要，可对这些对象进行重新映射。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对活动工作表中第一个 OLE 对象的 AutoLoad 属性进行设置。*/ Application.ActiveSheet.OLEObjects(1).AutoLoad = false` |

#### **OLEObjects.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **OLEObjects** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **OLEObjects.Enabled**

如果启用对象，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Enabled**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Height**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Interior**

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Locked**

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

*express*   一个代表 **OLEObjects** 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

#### **OLEObjects.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Placement**

返回或设置一个包含 **XlPlacement**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPlacement%20%E6%9E%9A%E4%B8%BE.html)常量的 **Variant** 值，它代表对象附加到它所在的单元格的方式。

**语法**

**express.Placement**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.PrintObject**

如果打印文档时也打印指定对象，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintObject**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.ShapeRange**

返回一个 **ShapeRange** 对象，它代表指定的一个或多个对象。只读。

**语法**

**express.ShapeRange**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.SourceName**

返回或设置一个 **String** 值，它代表指定对象的链接源名称。

**语法**

**express.SourceName**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.Width**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **OLEObjects** 对象的变量。

#### **OLEObjects.ZOrder**

返回指定对象的 z-次序位置。**Long** 型，只读。

**语法**

**express.ZOrder**

*express*   一个代表 **OLEObjects** 对象的变量。

**说明**

在任何对象集合中，z-次序尾端的对象为 *collection*(1)，z-次序前端的对象为 *collection*(*collection*.**Count**)。例如，如果活动工作表中有嵌入图表，z-次序尾端的图表为 `ActiveSheet.ChartObjects(1)`，z-次序前端的图表为 `ActiveSheet.ChartObjects(ActiveSheet.ChartObjects.Count)`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Sheet1 上嵌入的第一张图表的 z-次序位置。*/ alert("The chart's z-order position is " + Application.Worksheets.Item("Sheet1").ChartObjects(1).ZOrder` |

适用环境：web

适用平台：windows/linux