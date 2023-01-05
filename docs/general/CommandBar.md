#### **CommandBar**



代表容器应用程序中的一个命令栏。**CommandBar** 对象是 **CommandBars** 集合的成员。

**说明**

使用 **CommandBars**(*index*) 可返回一个 **CommandBar** 对象，其中 *index* 是命令栏的名称或索引号。下面的示例遍历命令栏集合以查找名为“Forms”的命令栏。如果找到，则本示例将显示该命令栏并保护其停靠状态。在本示例中，变量 **cb** 代表一个 **CommandBar** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   	let foundFlag = false  let cmdBars = Application.ActiveDocument.CommandBars    	for(let i=1;i <= cmdBars.Count;i++)  	{       		let cb = cmdBars.Item(i)       		if(cb.Name == "Forms")  		{           			cb.Protection = Application.Enum.msoBarNoChangeDock 			cb.Visible = true            			foundFlag = true         		}   	}   	if(!foundFlag)  	{      	alert("The collection does not contain a Forms command bar.")       } } ` |

可以使用名称或索引号在容器应用程序的可用菜单栏和工具栏列表中指定菜单栏或工具栏，但必须使用名称指定菜单、快捷菜单或子菜单（所有这些内容均由 **CommandBar** 对象代表）。本示例在**“Tools”**菜单底部添加一个新的菜单项。单击该新菜单项将运行名为“qtrReport”的过程。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   	let newItem = Application.ActiveDocument.CommandBars.Item("Tools").Controls.Add(Application.Enum.msoControlButton)   	newItem.BeginGroup = true    	newItem.Caption = "Make Report"  　　　　 newItem.FaceID = 0   	newItem.OnAction = "qtrReport" }` |

如果两个或两个以上的自定义菜单或子菜单同名，**CommandBars(index)** 将返回第一个具有该名称的菜单或子菜单。为了确保返回正确的菜单或子菜单，可找到显示该菜单的弹出式控件，然后对该弹出式控件应用 **CommandBar** 属性以返回代表该菜单的命令栏。假定在名为“Custom Tools”的工具栏中，第三个控件是一个弹出式控件，则本示例可将**“Save”**命令添加到该菜单的底部。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let viewMenu = Application.ActiveDocument.CommandBars.Item("Custom Tools").Controls.Item(3) 	viewMenu.Controls.Add(null,3)    //ID of Save command is 3 }` |

**方法**

|                                                              | 名称            | 说明                                                       |
| ------------------------------------------------------------ | --------------- | ---------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**      | 从集合中删除 **CommandBar** 对象。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindControl** | 获取一个符合指定条件的 **CommandBarControl** 对象。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**       | 将内置命令栏重置为其默认配置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowPopup**   | 将指定的命令栏作为快捷菜单，在指定坐标或当前光标位置显示。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdaptiveMenu** | 获取一个 **Boolean** 类型的值，该值指定命令栏是否应包含自适应菜单。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 获取一个 **Application** 对象，代表 **CommandBar** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**      | 如果指定的命令栏是容器应用程序的内置命令栏，则获取 **True**。如果是自定义命令栏，则返回 **False**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Context**      | 获取或设置一个可确定命令栏保存位置的字符串。该字符串由应用程序定义和解释。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Controls**     | 获取一个代表命令栏上的所有控件的 **CommandBarControls** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 获取一个 32 位整数，指示创建 **CommandBar** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**      | 获取或设置用于指定是否启用了指定 **CommandBar** 的 **Boolean** 值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**       | 获取或设置 **CommandBar** 的高度。可读写。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**        | 获取一个 **Long** 类型的值，该值代表集合中 **CommandBar** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**         | 设置或获取从对象左边缘算起 **CommandBar** 相对于屏幕的水平距离（以像素为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**         | 获取内置的 **CommandBar** 对象的名称。只读。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameLocal**    | 获取以容器应用程序的语言版本显示的内置命令栏名称，或者返回或设置自定义命令栏的名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 获取 **CommandBar** 对象的 **Parent** 对象。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**     | 获取或设置一个代表命令栏位置的 **MsoBarPosition** 常量。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowIndex**     | 获取或设置一个命令栏相对于同一停靠区域中其他命令栏的停靠顺序。该属性值可以是大于零的整数，也可以是以下 **MsoBarRow** 常量之一：**msoBarRowFirst** 或 **msoBarRowLast**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**          | 设置或获取指定的命令栏顶边到屏幕顶边的距离。对于固定命令栏，此属性返回或设置从命令栏到停靠区域顶部的距离。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**         | 获取命令栏的类型。只读。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**      | 获取或设置命令栏的 **Visible** 属性。如果命令栏可见，则为 **true**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**        | 获取或设置指定命令栏的宽度（以像素为单位）。可读写。         |

**成员方法**

#### **CommandBar.Delete**

从集合中删除 **CommandBar** 对象。

**语法**

**express.Delete()**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

对于 **Scripts** 集合，使用 **Delete** 方法将删除指定的 WPS 文档、ET 工作表或 WPP 幻灯片中的所有脚本。在宿主应用程序中，脚本定位标记以形状表示。因此，与每个 **msoScriptAnchor** 类型的脚本定位标记相关联的 **Shape** 对象将从 ET 和 WPP 的 **Shapes** 集合以及 WPS 的 **InlineShapes** 和 **Shapes** 集合中删除。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例删除所有不可见的自定义命令栏。 function test() {   	let foundFlag = false  	let delBars = 0 	let cmdBars = Application.ActiveDocument.CommandBars; 	for (let bar = 1; bar <= cmdBars.Count; bar++) { 		if (cmdBars.Item(bar).BuiltIn == false  	        && cmdBars.Item(bar).Visible == false) { 				 	        cmdBars.Item(bar).Delete() 	        foundFlag = true  	        delBars = delBars + 1 	    } 	} 	if (!foundFlag) {    	 alert("No command bars have been deleted.") 	} 	else {    	 alert(delBars + " custom bar(s) deleted.") 	} } ` |

#### **CommandBar.FindControl**

获取一个符合指定条件的 **CommandBarControl** 对象。

**语法**

**express.FindControl(Type, Id, Tag, Visible, Recursive)**

*express*   一个代表 **CommandBar** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*      | 可选          | **Variant**  | 控件的类型。                                                 |
| *Id*        | 可选          | **Variant**  | 控件的标识符。                                               |
| *Tag*       | 可选          | **Variant**  | 控件的标记值。                                               |
| *Visible*   | 可选          | **Variant**  | 如果为 true，则只能在搜索中包含可见的命令栏控件。默认值为 false。可见的命令栏包括所有可见的工具栏和执行 FindControl 方法时打开的所有菜单。 |
| *Recursive* | 可选          | **Variant**  | 如果该值为 true，那么将在命令栏及其所有弹出式子工具栏中查找。此参数仅应用于 CommandBar 对象。默认值为 false。 |

**说明**

如果 **CommandBars** 集合中有两个或者更多的控件符合搜索条件，那么 FindControl 返回找到的第一个控件。如果没有控件符合搜索条件，那么 **FindControl** 返回 Nothing。

#### **CommandBar.Reset**

将内置命令栏重置为其默认配置。

**语法**

**express.Reset()**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。重置一个内置命令栏将删除其中的自定义控件并恢复其内置控件。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例用 User 值来根据用户级别调整命令栏。如果 User 为“Level 1”，将显示命令栏“Custom”。如果 User 为任何其他值，那么将重置为默认状态并且命令栏“Custom”被禁用。 function test() {     let myBar = Application.CommandBars.Item("Custom")     if (user == "Level 1" )     {         myBar.Visible = true     }     else  	{         Application.CommandBars.Item("Visual Basic").Reset()         myBar.Enabled = false      } }` |

#### **CommandBar.ShowPopup**

将指定的命令栏作为快捷菜单，在指定坐标或当前光标位置显示。

**语法**

**express.ShowPopup(x, y)**

*express*   一个代表 **CommandBar** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *x*      | 可选          | **Variant**  | 快捷菜单所在位置的 x 坐标。如果省略此参数，那么将取当前光标位置的 x 坐标。 |
| *y*      | 可选          | **Variant**  | 快捷菜单所在位置的 y 坐标。如果省略此参数，那么将取当前光标位置的 y 坐标。 |

**说明**

当菜单左对齐时，**ShowPopup** 方法显示的快捷菜单的左上角位于 (x, y + 1)；当菜单右对齐时，快捷菜单的右上角位于 (x + 1, y + 1)。您可以使用 Windows 函数 **GetSystemMetrics(SM_MENUDROPALIGNMENT)** 检查下拉菜单对齐的系统布置。

当 (x, y) 坐标的屏幕位置导致全部或部分弹出式菜单超出可视屏幕的边界，弹出式菜单将发生平移，使其显示在可视区域中。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 如果命令栏的 **Position** 属性没有设置为 **msoBarPopup**，则此方法失败。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例可实现的功能为：创建一个包含两个控件的快捷菜单。ShowPopup 方法用于显示该快捷菜单。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarPopup, null, false)     myBar.Controls.Add(Application.Enum.msoControlButton, 3)     myBar.Controls.Add(Application.Enum.msoControlComboBox)     myBar.ShowPopup() }` |

**成员属性**

#### **CommandBar.AdaptiveMenu**

获取一个 **Boolean** 类型的值，该值指定命令栏是否应包含自适应菜单。可读写。

**语法**

**express.AdaptiveMenu**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBar.Application**

获取一个 **Application** 对象，代表 **CommandBar** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBar.BuiltIn**

如果指定的命令栏是容器应用程序的内置命令栏，则获取 **True**。如果是自定义命令栏，则返回 **False**。只读。

**语法**

**express.BuiltIn**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例删除所有未在屏幕上显示的自定义命令栏。 function test() {     let foundFlag = false      let deletedBars = 0     for(let i=1;i <= Application.CommandBars.Count;i++)      {         let bar = Application.CommandBars.Item(i)         if(bar.BuiltIn == false && bar.Visible == false)      	{             bar.Delete()             foundFlag = true              deletedBars++         }     }     if(!foundFlag)      {         alert("No command bars have been deleted.")     }     else      {         alert(deletedBars + " custom command bar(s) deleted.")     } }` |

#### **CommandBar.Context**

获取或设置一个可确定命令栏保存位置的字符串。该字符串由应用程序定义和解释。可读写。

**语法**

**express.Context**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一个消息框，该消息框包含命令栏“Custom”的上下文相关字符串。本示例运行于 WPS 和其他支持 Context 属性的应用程序中。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     myBar.Controls.Add(Application.Enum.msoControlButton, 2)     myBar.Visible = true      alert(myBar.Context) }` |

#### **CommandBar.Controls**

获取一个代表命令栏上的所有控件的 **CommandBarControls** 对象。只读。

**语法**

**express.Controls**

*express*   一个代表 **CommandBar** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在命令栏“Custom”中添加一个组合框控件，然后在该列表中添充两项。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。*/ function test() {   	let myControl = Application.ActiveDocument.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, null, null, 1) 	myControl.AddItem("First Item", 1) 	myControl.AddItem("Second Item", 2) 	myControl.DropDownLines = 3 	myControl.DropDownWidth = 75 	myControl.ListHeaderCount = 0 } ` |

#### **CommandBar.Creator**

获取一个 32 位整数，指示创建 **CommandBar** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBar.Enabled**

获取或设置用于指定是否启用了指定 **CommandBar** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **t****rue**，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **true**，则将由应用程序确定其状态。但如果将该属性设置为 **false**，则强行禁用该控件。

#### **CommandBar.Height**

获取或设置 **CommandBar** 的高度。可读写。

**语法**

**express.Height**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置该控件宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。 function test() {     let myBar = Application.CommandBars.Item("Custom")     let barHeight = myBar.Height     let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)     myControl.Height = barHeight * 2     myControl.Width = 50     myBar.Visible = true }` |

#### **CommandBar.Index**

获取一个 **Long** 类型的值，该值代表集合中 **CommandBar** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBar.Left**

设置或获取从对象左边缘算起 **CommandBar** 相对于屏幕的水平距离（以像素为单位）。可读写。

**语法**

**express.Left**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将名为 Custom 的命令栏从其固定位置沿窗口顶部移动窗口的左边。 function test() {     let myBar = Application.CommandBars.Item("Custom")     myBar.Position = 1     myBar.RowIndex = 2     myBar.Left = 0 }` |

#### **CommandBar.Name**

获取内置的 **CommandBar** 对象的名称。只读。

**语法**

**express.Name**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

内置命令栏的本地名称显示在标题栏中（在命令栏未固定时）以及可用命令栏列表中（无论该列表显示在容器应用程序的什么位置）。对于内置命令栏，**Name** 属性返回该命令栏的美式英语名称。使用 **NameLocal** 属性可返回其本地化的名称。如果更改了一个自定义命令栏的 **LocalName** 属性值，则该命令栏的 **Name** 值也将发生更改，反之亦然。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例可实现的功能为：在命令栏集合中查找命令栏“Custom”，如果找到，那么本示例显示该命令栏。 function test() {     let foundFlag =  false     for(let i=1;i <= Application.CommandBars.Count;i++)     {         let bar = Application.CommandBars.Item(i)         if(bar.Name == "Custom")     	{             foundFlag = true             bar.Visible = true         }     }     if(!foundFlag)     {         alert("'Custom' bar isn't in collection.")     }     else     {         alert("'Custom' bar is now visible.")     } }` |

#### **CommandBar.NameLocal**

获取以容器应用程序的语言版本显示的内置命令栏名称，或者返回或设置自定义命令栏的名称。可读写。

**语法**

**express.NameLocal**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

无论有效命令栏列表显示在容器应用程序的哪个位置，内置命令栏的本地名都将显示在标题栏中（当该命令栏未固定时），以及有效命令栏列表中。

如果更改了一个自定义命令栏的 **LocalName** 属性值，则该命令栏的 **Name** 值也将发生更改，反之亦然。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| ![img]()注释                                 |
| -------------------------------------------- |
| 如果要对内置命令栏设置该属性，就会出现错误。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例可实现的功能为：显示容器应用程序中第一个命令栏的名称和本地化名称。  function test() {     let Com = Application.CommandBars.Item(1)     alert("The name of the command bar is " + Com.Name)     alert("The localized name of the command bar is " + Com.NameLocal) } ` |

#### **CommandBar.Parent**

获取 **CommandBar** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBar.Position**

获取或设置一个代表命令栏位置的 **MsoBarPosition** 常量。可读写。

**语法**

**express.Position**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对命令栏集合中的元素逐个进行如下操作：将自定义命令栏停靠在应用程序窗口的底部，将内置命令栏停靠在窗口的顶部。   function test() {     for(let i=1;i <= Application.CommandBars.Count;i++)  	{         let bar = Application.CommandBars.Item(i)         if(bar.Visible == true)  	    {             if(bar.BuiltIn) {                 bar.Position = Application.Enum.msoBarTop             }             else {                 bar.Position = Application.Enum.msoBarBottom             }         }     } }` |

#### **CommandBar.RowIndex**

获取或设置一个命令栏相对于同一停靠区域中其他命令栏的停靠顺序。该属性值可以是大于零的整数，也可以是以下 **MsoBarRow** 常量之一：**msoBarRowFirst** 或 **msoBarRowLast**。可读写。

**语法**

**express.RowIndex**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

多个命令栏可享有同一行索引号，具有较低数值的命令栏先固定。如果存在两个或两个以上命令栏具有同一行索引号的情况，那么最后指定的命令栏在这一组中最先显示。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将名为“Custom”的命令栏的位置调整到默认值以左 110 个像素的位置，并通过将其行索引号改为 msoBarRowFirst 使该命令栏最先固定。 function test() {     let myBar = Application.CommandBars.Item("Custom")     myBar.RowIndex = Application.Enum.msoBarRowFirst     myBar.Left = 140 }` |

#### **CommandBar.Top**

设置或获取指定的命令栏顶边到屏幕顶边的距离。对于固定命令栏，此属性返回或设置从命令栏到停靠区域顶部的距离。可读写。

**语法**

**express.Top**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将名为“Custom”的浮动命令栏的左上角定位为距离屏幕左边 140 像素、距离屏幕顶部 100 像素。 function test() {     let myBar = Application.CommandBars.Item("Custom")     myBar.Position = Application.Enum.msoBarFloating     myBar.Left = 140     myBar.Top = 100 }` |

#### **CommandBar.Type**

获取命令栏的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例查找名为“Custom”的命令栏上的第一个控件，并使用 Type 属性确定该控件是否为按钮。如果该控件为按钮，本示例将复制“Copy”（位于“Standard”工具栏上）按钮的图标并将其粘贴到该控件上。 function test() {     let oldCtrl = Application.CommandBars.Item("Custom").Controls.Item(1)     if(oldCtrl.Type == Application.Enum.msoControlButton)      {         let newCtrl = Application.CommandBars.FindControl(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Copy").ID)         newCtrl.CopyFace()         oldCtrl.PasteFace()     } }` |

#### **CommandBar.Visible**

获取或设置命令栏的 **Visible** 属性。如果命令栏可见，则为 **true**。可读写。

**语法**

**express.Visible**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

默认情况下，新创建的自定义命令栏的 **Visible** 属性为 **false**。

必须先将命令栏的 **Enabled** 属性设置为 **true**，才能将其 **Visible** 属性设置为 **true**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将遍历命令栏集合，以查找“Forms”命令栏。如果找到“Forms”命令栏，则显示该命令栏并保护其停靠状态。 */  function test() {   	let foundFlag = false  	for(let i=1;i <= Application.CommandBars.Count;i++) {     let cmdbar = Application.CommandBars.Item(i)     if(cmdbar.Name == "Forms") {         cmdbar.Protection = msoBarNoChangeDock         cmdbar.Visible = true          foundFlag = true      } 	} 	if(!foundFlag) {     	alert("'Forms'command bar is not in the collection.") 	} } ` |

#### **CommandBar.Width**

获取或设置指定命令栏的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

*express*   一个代表 **CommandBar** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。 function test() {     let myBar = Application.CommandBars.Item("Custom")     let barHeight = myBar.Height     let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)     myControl.Height = barHeight * 2     myControl.Width = 50     myBar.Visible = true }` |

适用环境：web

适用平台：windows/linux