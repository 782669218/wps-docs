#### **CommandBars**



代表容器应用程序中命令栏的 **CommandBar** 对象的集合。

**说明**

在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。

使用 **CommandBars** 属性可返回 **CommandBars** 集合。下面的示例在**“立即”**窗口中显示每个菜单栏和工具栏的名称和本地名称，并显示一个值以指示该菜单栏或工具栏是否可见。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	for(let i=1;i<=CommandBars.Count;i++){     	let cbar = CommandBars.Item(i)     	alert(cbar.Name + "\t" + cbar.NameLocal + "\t" + cbar.Visible) 	}	 }` |

使用 **Add** 方法可向集合中添加一个新的命令栏。下面的示例创建一个名为“Custom1”的自定义工具栏并将其显示为浮动工具栏。  

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let cbar1 = Application.ActiveDocument.CommandBars.Add("Custom1", msoBarFloating) 	cbar1.Visible = true }` |

使用 enumName 可返回一个 **CommandBar** 对象，其中 *index* 是该命令栏的名称或索引号。下面的示例将名为“Custom1”的工具栏停靠在应用程序窗口的底部。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.CommandBars.Item("Custom1").Position = msoBarBottom` |

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释         |
| ------------------------------------------------------------ |
| 可以使用名称或索引号指定容器应用程序的可用菜单栏和工具栏列表中的菜单栏或工具栏，但必须使用名称指定菜单、快捷菜单或子菜单（所有这些均由 **CommandBar** 对象代表）。如果两个或两个以上的自定义菜单或子菜单同名，则 enumName 将返回第一个具有该名称的菜单。为确保返回正确的菜单或子菜单，可以找到显示该菜单的弹出式控件，然后对该弹出式控件应用 **CommandBar** 属性以返回代表该菜单的命令栏。 |

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                        | 创建一个新的命令栏并将其添加到命令栏集合中。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CommitRenderingTransaction** | 提交呈现事务                                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExecuteMso**                 | 执行由 **idMso** 参数标识的控件。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindControl**                | 获取一个符合指定条件的 **CommandBarControl** 对象。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindControls**               | 获取符合指定条件的 **CommandBarControls** 集合。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetEnabledMso**              | 如果启用了 **idMso** 参数标识的控件，则返回 True。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetImageMso**                | 返回由 **idMso** 参数标识的控件图像（缩放到指定的宽度和高度尺寸）的 **IPictureDisp** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetLabelMso**                | 将 **idMso** 参数标识的控件的标签作为 String 类型的值返回。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetPressedMso**              | 返回一个值，该值指示是否已按下 **idMso** 参数标识的 toggleButton 控件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetScreentipMso**            | 将 **idMso** 参数标识的控件的屏幕提示作为 String 类型的值返回。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetSupertipMso**             | 将 **idMso** 参数标识的控件的超级提示作为 String 类型的值返回。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetVisibleMso**              | 如果 **idMso** 参数标识的控件可见，则返回 true。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReleaseFocus**               | 释放所有命令栏的用户界面焦点。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetVisibleMso**              | 通过控件对应的idMso控制该控件的显示状态(显示/隐藏）。        |

**属性**

|                                                              | 名称                            | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActionControl**               | 获取一个 **CommandBarControl** 对象，该对象的 **OnAction** 属性设置为当前正在运行的过程。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveMenuBar**               | 获取一个 **CommandBar** 对象，该对象代表容器应用程序中的活动菜单栏。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdaptiveMenus**               | 此属性可选中或取消选中指定 WPS Office中菜单是完全显示还是按个性化方式显示的选项的复选框控件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                 | 获取一个 **Application** 对象，代表 **CommandBars** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                       | 获取宿主应用程序中的命令栏数量。只读。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                     | 获取一个 32 位整数，指示创建 **CommandBars** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableAskAQuestionDropdown** | 如果启用了**“应答向导”**下拉菜单，则为 **True**。可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableCustomize**            | 如果禁用了工具栏的自定义功能，则为 **True**。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayFonts**                | 如果在**“字体”**框中以实际字体显示字体名称，则为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayKeysInTooltips**       | 如果每个命令栏控件的快捷键都显示在**“工具提示”**中，则为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayTooltips**             | 如果只要用户将指针放在命令栏控件上方就显示“屏幕提示”，则为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**                        | 获取 **CommandBars** 集合中的 **CommandBar** 对象。只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LargeButtons**                | 如果显示的工具栏按钮比常规尺寸要大，则为 **True**。可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MenuAnimationStyle**          | 获取或设置一个代表命令栏的动画方式的 **MsoMenuAnimation**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                      | 获取 **CommandBars** 对象的 **Parent** 对象。只读。          |

**成员方法**

#### **CommandBars.Add**

创建一个新的命令栏并将其添加到命令栏集合中。

**语法**

**express.Add(Name, Position, MenuBar, Temporary)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*      | 可选          | **Variant**  | 新命令栏的名称。如果省略此参数，则为命令栏指定默认名称（如 Custom 1）。 |
| *Position*  | 可选          | **Variant**  | 新命令栏的位置或类型。可以是 MsoBarPosition 常量之一。       |
| *MenuBar*   | 可选          | **Variant**  | 设置为 true 将以新命令栏替换活动菜单栏。默认值为 false。     |
| *Temporary* | 可选          | **Variant**  | 如果为 true，则将使新命令栏变成临时命令栏。临时命令栏将在容器应用程序关闭时删除。默认值为 false。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let cbar1 = Application.ActiveDocument.CommandBars.Add("Custom1", msoBarFloating) 	cbar1.Visible = true }` |

#### **CommandBars.CommitRenderingTransaction**

提交呈现事务

**语法**

**express.CommitRenderingTransaction(hwnd)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *hwnd*   | 必选          | **Long**     | 要在其中提交呈现事务的窗口的句柄。 |

#### **CommandBars.ExecuteMso**

执行由 **idMso** 参数标识的控件。

**语法**

**express.ExecuteMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**说明**

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释   |
| ------------------------------------------------------ |
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

在某个命令没有对象模型的情况下，可以使用此方法。此方法可用于内置按钮、toggleButton 和 splitButton 类型的控件。此方法执行失败时将返回 E_InvalidArg（如果 **IdMso** 参数无效）或 E_Fail（若未启用控件或控件不可见）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例执行“Copy”按钮。 Application.ActiveDocument.CommandBars.ExecuteMso("Copy")` |

#### **CommandBars.FindControl**

获取一个符合指定条件的 **CommandBarControl** 对象。

**语法**

**express.FindControl(Type, Id, Tag, Visible)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*    | 可选          | **Variant**  | 控件的类型。                                                 |
| *Id*      | 可选          | **Variant**  | 控件的标识符。                                               |
| *Tag*     | 可选          | **Variant**  | 控件的标记值。                                               |
| *Visible* | 可选          | **Variant**  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。可见的命令栏包括所有可见的工具栏和执行 FindControl 方法时打开的所有菜单。 |

**说明**

如果 **CommandBars** 集合中有两个或者更多的控件符合搜索条件，那么 FindControl 返回找到的第一个控件。如果没有控件符合搜索条件，那么 **FindControl** 返回 Nothing。

#### **CommandBars.FindControls**

获取符合指定条件的 **CommandBarControls** 集合。

**语法**

**express.FindControls(Type, Id, Tag, Visible)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*    | 可选          | **Variant**  | 为指定控件类型的 MsoControlType 常量之一。                   |
| *Id*      | 可选          | **Variant**  | 控件的标识符。                                               |
| *Tag*     | 可选          | **Variant**  | 控件的标记值。                                               |
| *Visible* | 可选          | **Variant**  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。 |

**说明**

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)             |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

如果未找到符合条件的控件，**FindControls** 方法将返回 **Nothing**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例用 FindControls 方法返回 CommandBars 集合中所有 ID 为 18 的成员，并在消息框中显示符合指定条件的控件的个数。 function test() { 	let  myControls = CommandBars.FindControls(msoControlButton,18) 	alert("There are " + myControls.Count + " controls that meet the search criteria.") }` |

#### **CommandBars.GetEnabledMso**

如果启用了 **idMso** 参数标识的控件，则返回 True。

**语法**

**express.GetEnabledMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例在“加粗”按钮已启用的情况下返回 True。 Application.CommandBars.GetEnabledMso("Bold")` |

#### **CommandBars.GetImageMso**

返回由 **idMso** 参数标识的控件图像（缩放到指定的宽度和高度尺寸）的 **IPictureDisp** 对象。

**语法**

**express.GetImageMso(idMso, Width, Height)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |
| *Width*  | 必选          | **Number**   | 图像的宽度     |
| *Height* | 必选          | **Number**   | 图像的高度。   |

**说明**

**Width** 和 **Height** 参数的值必须介于 16 和 128 之间。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例将 32x32 尺寸版本的“Paste”图标作为 IPictureDisp 对象返回。 Application.CommandBars.GetImageMso("Paste", 32, 32)` |

#### **CommandBars.GetLabelMso**

将 **idMso** 参数标识的控件的标签作为 String 类型的值返回。

**语法**

**express.GetLabelMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回 String 类型的值“Paste”。  Application.CommandBars.GetLabelMso("Paste")` |

#### **CommandBars.GetPressedMso**

返回一个值，该值指示是否已按下 **idMso** 参数标识的 toggleButton 控件。

**语法**

**express.GetPressedMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**返回值**

Boolean

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例在“加粗”按钮已按下的情况下返回 True。 Application.CommandBars.GetPressedMso("Bold") ` |

#### **CommandBars.GetScreentipMso**

将 **idMso** 参数标识的控件的屏幕提示作为 String 类型的值返回。

**语法**

**express.GetScreentipMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**返回值**

String

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回 String 类型的值“Paste”。 Application.CommandBars.GetScreentipMso("Paste")` |

#### **CommandBars.GetSupertipMso**

将 **idMso** 参数标识的控件的超级提示作为 String 类型的值返回。

**语法**

**express.GetSupertipMso(idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回 String 类型的值“从文档中剪切所选内容，并将其放入剪贴板”。 Application.CommandBars.GetSupertipMso("Cut")` |

#### **CommandBars.GetVisibleMso**

如果 **idMso** 参数标识的控件可见，则返回 true。

**语法**

**express.GetVisibleMso (idMso)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *idMso*  | 必选          | **String**   | 控件的标识符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例在“加粗”按钮可见的情况下返回 True。 Application.CommandBars.GetVisibleMso("Bold")` |

#### **CommandBars.ReleaseFocus**

释放所有命令栏的用户界面焦点。

**语法**

**express.ReleaseFocus()**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向命令栏“Custom”中添加三个空白按钮，并将焦点设置在中间的按钮上。等待五秒钟后，释放所有命令栏的用户界面焦点。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined,true)     myBar.Controls.Add(Application.Enum.msoControlButton)     myBar.Controls.Add(Application.Enum.msoControlButton)     myBar.Controls.Add(Application.Enum.msoControlButton)     myBar.Visible = true      let myControl = Application.CommandBars.Item("Custom").Controls.Item(2)     myControl.SetFocus()     let PauseTime = 5   // Set duration.     let Start = Timer   // Set start time.     while(Timer(Start + PauseTime)) 	{         DoEvents()          // Yield to other processes.         let Finish = Timer         Application.CommandBars.ReleaseFocus()     } }` |

#### **CommandBars.SetVisibleMso**

通过控件对应的idMso控制该控件的显示状态(显示/隐藏）。

**语法**

**express.SetVisibleMso(idMso, Enabled)**

*express*   一个代表 **CommandBars** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**       |
| --------- | ------------- | ------------ | -------------- |
| *idMso*   | 必选          | **String**   | 控件的标识符。 |
| *Enabled* | 必选          | **Boolean**  | 是否隐藏。     |

**成员属性**

#### **CommandBars.ActionControl**

获取一个 **CommandBarControl** 对象，该对象的 **OnAction** 属性设置为当前正在运行的过程。只读。

**语法**

**express.ActionControl**

*express*   一个代表 **CommandBars** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建命令栏“Custom”，向命令栏中添加三个按钮，然后用 ActionControl 属性和 Tag 属性确定最后一次单击的是哪一个命令栏按钮。 function test() { 	let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined,true) 	let buttonOne = myBar.Controls.Add(Application.Enum.msoControlButton)     buttonOne.FaceId = 133     buttonOne.Tag = "RightArrow"     buttonOne.OnAction = "whichButton" 	let buttonTwo = myBar.Controls.Add(Application.Enum.msoControlButton)     buttonTwo.FaceId = 134     buttonTwo.Tag = "UpArrow"     buttonTwo.OnAction = "whichButton" 	let buttonThree = myBar.Controls.Add(Application.Enum.msoControlButton)     buttonThree.FaceId = 135     buttonThree.Tag = "DownArrow"     buttonThree.OnAction = "whichButton" 	myBar.Visible = true } //下面的子例程响应 OnAction 方法并确定最后单击的是哪一个命令栏按钮。  function whichButton()  {     switch(Application.ActiveDocument.CommandBars.ActionControl.Tag) {         case "RightArrow":             alert("Right Arrow button clicked.")             break         case "UpArrow":             alert ("Up Arrow button clicked.")             break         case "DownArrow":             alert ("Down Arrow button clicked.")             break     }  }` |

#### **CommandBars.ActiveMenuBar**

获取一个 **CommandBar** 对象，该对象代表容器应用程序中的活动菜单栏。只读。

**语法**

**express.ActiveMenuBar**

*express*   一个代表 **CommandBars** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在活动菜单栏的末尾暂时添加一个弹出式控件“Custom”，然后向该弹出式控件中添加一个名为“Import”的控件。 */ function test() { 	let myMenuBar = Application.ActiveDocument.CommandBars.ActiveMenuBar 	let newMenu = myMenuBar.Controls.Add(Application.Enum.msoControlPopup,undefined,undefined,undefined,true)     newMenu.Caption = "Custom" 	let ctrl1 = newMenu.CommandBar.Controls.Add(Application.Enum.msoControlButton,1)     ctrl1.Caption = "Import"     ctrl1.TooltipText = "Import"     ctrl1.Style = Application.Enum.msoButtonCaption }` |

#### **CommandBars.AdaptiveMenus**

此属性可选中或取消选中指定 WPS Office中菜单是完全显示还是按个性化方式显示的选项的复选框控件。可读写。

**语法**

**express.AdaptiveMenus**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

如果对所有 WPS Office应用程序都启用了自适应菜单，则为 **True**。可读写 **Boolean** 类型。

在任意应用程序中执行下列操作均可设置此控件：

1. 在**“工具”**菜单上，选择**“自定义”**。
2. 选择**“选项”**选项卡。
3. **“始终显示整个菜单”**选项位于**“个性化菜单和工具栏”**部分中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 WPS Office中的所有命令栏（其中包括自定义命令栏及其上的自定义命令控件）设置了三个选项。 Application.ActiveDocument.CommandBars.LargeButtons = true  Application.ActiveDocument.CommandBars.DisplayFonts = true Application.ActiveDocument.CommandBars.AdaptiveMenus = true` |

#### **CommandBars.Application**

获取一个 **Application** 对象，代表 **CommandBars** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBars.Count**

获取宿主应用程序中的命令栏数量。只读。

**语法**

**express.Count**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBars.Creator**

获取一个 32 位整数，指示创建 **CommandBars** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBars.DisableAskAQuestionDropdown**

如果启用了**“应答向导”**下拉菜单，则为 **True**。可读写。

**语法**

**express.DisableAskAQuestionDropdown**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例打开或关闭 DisableAskAQuestionDropdown 属性。 function test() {     let comBars = Application.CommandBars     if(comBars.DisableAskAQuestionDropdown == true) {         comBars.DisableAskAQuestionDropdown = false      }     else {         comBars.DisableAskAQuestionDropdown = true      }  }` |

#### **CommandBars.DisableCustomize**

如果禁用了工具栏的自定义功能，则为 **True**。可读写。

**语法**

**express.DisableCustomize**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例打开或关闭 DisableCustomize 属性。 function test() {     let comBars = Application.CommandBars     if(comBars.DisableCustomize == true) 	{         comBars.DisableCustomize = false     }     else{         comBars.DisableCustomize = true     }  }` |

#### **CommandBars.DisplayFonts**

如果在**“字体”**框中以实际字体显示字体名称，则为 **True**。可读写。

**语法**

**express.DisplayFonts**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 WPS Office中的所有命令栏（其中包括自定义命令栏及其上的自定义命令控件）设置了三个选项。  function test() {     Application.CommandBars.LargeButtons = true      Application.CommandBars.DisplayFonts = true     Application.CommandBars.AdaptiveMenus = true  }` |

#### **CommandBars.DisplayKeysInTooltips**

如果每个命令栏控件的快捷键都显示在**“工具提示”**中，则为 **True**。可读写。

**语法**

**express.DisplayKeysInTooltips**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

要在**“工具提示”**中显示快捷键，还必须将 **DisplayTooltips** 属性设置为 **True**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 WPS Office中所有命令栏设置选项。  function test() {     Application.CommandBars.LargeButtons = true     Application.CommandBars.DisplayTooltips = true      Application.CommandBars.DisplayKeysInTooltips = true     Application.CommandBars.MenuAnimationStyle = Application.Enum.msoMenuAnimationUnfold  }` |

#### **CommandBars.DisplayTooltips**

如果只要用户将指针放在命令栏控件上方就显示“屏幕提示”，则为 **True**。可读写。

**语法**

**express.DisplayTooltips**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

在容器应用程序中设置 **DisplayTooltips** 属性会立刻影响当时正在运行的每个 WPS Office应用程序中的所有命令栏和设置该属性后打开的所有 Office 应用程序中的所有命令栏。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在所有命令栏上显示大控件和工具提示。   function test() {     Application.CommandBars.LargeButtons = true      Application.CommandBars.DisplayTooltips = true }` |

#### **CommandBars.Item**

获取 **CommandBars** 集合中的 **CommandBar** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **CommandBars** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//Item 为对象或集合的默认成员。下面语句将 CommandBar 对象指定给 cmdBar。 let cmdBar = Application.ActiveDocument.CommandBars.Item("Standard")` |

#### **CommandBars.LargeButtons**

如果显示的工具栏按钮比常规尺寸要大，则为 **True**。可读写。

**语法**

**express.LargeButtons**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在所有命令栏上的工具栏按钮的显示大小间切换。   function test() {     if(Application.CommandBars.LargeButtons)     {         Application.CommandBars.LargeButtons = false     }     else     {         Application.CommandBars.LargeButtons = true     } }` |

#### **CommandBars.MenuAnimationStyle**

获取或设置一个代表命令栏的动画方式的 **MsoMenuAnimation**。可读写。

**语法**

**express.MenuAnimationStyle**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为 WPS Office中所有命令栏设置选项。   function test() {     Application.CommandBars.LargeButtons = true     Application.CommandBars.DisplayTooltips = true      Application.CommandBars.DisplayKeysInTooltips = true      Application.CommandBars.MenuAnimationStyle = Application.Enum.msoMenuAnimationUnfold }` |

#### **CommandBars.Parent**

获取 **CommandBars** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CommandBars** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

适用环境：web

适用平台：windows/linux