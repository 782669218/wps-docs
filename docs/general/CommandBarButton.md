#### **CommandBarButton**



代表命令栏中的一个按钮控件。

**说明**

使用 **Controls(index)** 可返回一个 **CommandBarButton** 对象，其中 *index* 是控件的索引号。请注意，该控件的 **Type** 属性必须是 **msoControlButton**。假定在名为“Custom”的命令栏上，第二个控件是一个按钮，则下面的示例将更改该按钮的样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let cmdBars = Application.ActiveDocument.CommandBars 	let c = cmdBars.Item("Custom").Controls.Item(2) 	if ( c.Type == Application.Enum.msoControlButton ) {    		if ( c.Style == Application.Enum.msoButtonIcon ) { 			c.Style = Application.Enum.msoButtonIconAndCaption     	}     	else { 			c.Style = Application.Enum.msoButtonIcon    		} 	} }` |

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**     | 将一个命令栏按钮控件复制到已有的命令栏中。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyFace** | 将命令栏按钮控件的图标复制到“剪贴板”中。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**   | 将 **CommandBarButton** 对象从其集合中删除。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute**  | 运行分配给指定 **CommandBarButton** 控件的过程或内置命令。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**     | 将指定的 **CommandBarButton** 控件移动到已有的命令栏。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**    | 将内置 **CommandBarButton** 控件重置为其初始功能和图符。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetFocus** | 将键盘的焦点移到指定的 **CommandBarButton** 控件。如果该按钮被禁用或不可见，那么此方法将失败。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 获取一个 **Application** 对象，代表 **CommandBarButton** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginGroup**        | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**           | 如果指定的命令栏控件是容器应用程序的控件，则为 **True**。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltInFace**       | 如果命令栏按钮控件的图标是其原始内置图标，则为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**           | 获取或设置命令栏控件的标题文字。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 获取一个 32 位整数，指示创建 **CommandBarButton** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DescriptionText**   | 获取或设置命令栏按钮控件的说明。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**           | 如果启用了指定的 **CommandBar** 或 **CommandBarControl**，则此属性为 true。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FaceId**            | 获取或设置某个 **CommandBarButton** 控件的图符的 ID 号。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**            | 获取或设置命令栏控件的高度。可读写。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpContextId**     | 获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的帮助上下文 ID 号。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpFile**          | 获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的文件名。可读写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HyperlinkType**     | 设置或获取一个 **MsoCommandBarButtonHyperlinkType** 常量，代表与指定的命令栏按钮相关联的超链接的类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                | 获取内置的 **CommandBarButton** 控件的 ID。只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**             | 获取一个 **Long** 类型的值，该值代表集合中 **CommandBarButton** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsPriorityDropped** | 如果当前已根据使用统计信息和布局空间将 **CommandBarButton** 控件从菜单或工具栏中删除，则获取 **True**。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**              | 设置或获取指定的 **CommandBarButton** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Mask**              | 获取或设置代表 **CommandBarButton** 对象的屏蔽图像的 **IPictureDisp** 对象。屏蔽图像可决定按钮图像的透明部分。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEUsage**          | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarButton** 控件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OnAction**          | 获取或设置一个 JS函数的名称，该过程在用户单击或更改 **CommandBarButton** 控件的值时运行。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parameter**         | 获取或设置一个应用程序可用于执行 **CommandBarButton** 控件中命令的字符串。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 获取 **CommandBarButton** 对象的 **Parent** 对象。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Picture**           | 获取或设置一个代表 **CommandBarButton** 对象的图像的 **IPictureDisp** 对象。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**          | 获取或设置 CommandBarButton 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShortcutText**      | 获取或设置当控钮出现在菜单、子菜单或快捷菜单上时 **CommandBarButton** 控件旁边显示的快捷键文字。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **State**             | 获取或设置 CommandBarButton 控件的外观。可读写。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**             | 获取或设置 **CommandBarButton** 控件的显示方式。可读写。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tag**               | 获取或设置有关 **CommandBarButton** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TooltipText**       | 获取或设置 **CommandBarButton**的**“屏幕提示”**中显示的文本。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**               | 获取指定的 **CommandBarButton** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 获取 **CommandBarButton** 控件的类型。只读。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**           | 获取或设置 **CommandBarButton** 控件的 **Visible** 属性。如果 **CommandBarButton** 可见，则此属性为 **true**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**             | 获取或设置指定的 **CommandBarButton** 控件的宽度（以像素为单位）。可读写。 |

**成员方法**

#### **CommandBarButton.Copy**

将一个命令栏按钮控件复制到已有的命令栏中。

**语法**

**express.Copy(Bar, Before)**

*express*   一个代表 **CommandBarButton** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。 |
| *Before* | 可选          | **Variant**  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

#### **CommandBarButton.CopyFace**

将命令栏按钮控件的图标复制到“剪贴板”中。

**语法**

**express.CopyFace()**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

可使用 **PasteFace** 方法将“剪贴板”上的内容粘贴到一个按钮图标上。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarButton.Delete**

将 **CommandBarButton** 对象从其集合中删除。

**语法**

**express.Delete(Temporary)**

*express*   一个代表 **CommandBarButton** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Temporary* | 可选          | **Variant**  | 如果为 true，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

#### **CommandBarButton.Execute**

运行分配给指定 **CommandBarButton** 控件的过程或内置命令。

**语法**

**express.Execute()**

*express*   一个代表 **CommandBarButton** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET 标准命令格的第一个函数函数。` |

#### **CommandBarButton.Move**

将指定的 **CommandBarButton** 控件移动到已有的命令栏。

**语法**

**express.Move(Bar, Before)**

*express*   一个代表 **CommandBarButton** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选          | **Variant**  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。 */ function test() { 	let allControls = Application.ActiveDocument.CommandBars.Item("Custom").Controls 	for(let i=1; i<=allControls.Count; i++) {     if(allControls.Item(i).Type == Application.Enum.msoControlComboBox) {         allControls.Item(i).Move(null,7)         allControls.Item(i).Tag = "Selection box"         allControls.Item(i).Priority = 5     } } }` |

#### **CommandBarButton.Reset**

将内置 **CommandBarButton** 控件重置为其初始功能和图符。

**语法**

**express.Reset()**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例自定义一个命令栏按钮。首先，将按钮属性重置为其默认状态。然后设置各种按钮属性。 function test() {     let cbButton = Application.CommandBars.Item("Custom").Controls.Item(2)     cbButton.Reset()     cbButton.BuiltInFace = true      cbButton.Caption = "Compute Total"     cbButton.DescriptionText = "This button computes the total of all purchases."     cbButton.Enabled = true      cbButton.TooltipText = "Click to compute total amount for all items in your cart." }` |

#### **CommandBarButton.SetFocus**

将键盘的焦点移到指定的 **CommandBarButton** 控件。如果该按钮被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus()**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

带有焦点的控件与其他控件的差别是十分细微的。在使用此方法后，会看到该控件处于三维突出显示的状态。按方向键将使焦点在工具栏中各控件间切换，就好像是按键盘控制键到达该控件一样。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建一个名为“Custom”的命令栏，并在其中添加一个 ComboBox 控件和一个 Button 控件，然后使用 SetFocus 方法为 ComboBox 控件设置焦点。 function test() {     let focusBar = Application.CommandBars.Add("Custom")     Application.CommandBars.Item("Custom").Visible = true     Application.CommandBars.Item("Custom").Position = Application.Enum.msoBarTop     									     let testComboBox = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, 1)     									     let testButton = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlButton)     testButton.FaceId = 17     // Set the focus to the combo box.     testComboBox.SetFocus() }` |

**成员属性**

#### **CommandBarButton.Application**

获取一个 **Application** 对象，代表 **CommandBarButton** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.BeginGroup**

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。

**语法**

**express.BeginGroup**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.BuiltIn**

如果指定的命令栏控件是容器应用程序的控件，则为 **True**。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False**。只读。

**语法**

**express.BuiltIn**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.BuiltInFace**

如果命令栏按钮控件的图标是其原始内置图标，则为 **True**。可读写。

**语法**

**express.BuiltInFace**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

此属性只能设置为 **True**，其功能是将图标重新设置为内置的图标。可读写 **Boolean** 类型。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例可实现的功能为：确定命令栏“Custom”中第一个控件的图标是否为其内置的按钮图标。如果是，那么本示例将该按钮图标复制到“剪贴板”上。 function test() {     let myControl = Application.CommandBars.Item("Custom").Controls.Item(1)     if ( myControl.BuiltInFace == true)     {         myControl.CopyFace()     } }` |

#### **CommandBarButton.Caption**

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

*express*   一个代表 **CommandBarButton** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。*/  function test() { 	let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true) 	myBar.Visible = true  	let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2) 	myControl.DescriptionText = "Starts the spelling checker" 	myControl.Caption = "Spelling checker" }` |

#### **CommandBarButton.Creator**

获取一个 32 位整数，指示创建 **CommandBarButton** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.DescriptionText**

获取或设置命令栏按钮控件的说明。可读写。

**语法**

**express.DescriptionText**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

#### **CommandBarButton.Enabled**

如果启用了指定的 **CommandBar** 或 **CommandBarControl**，则此属性为 true。可读写。

**语法**

**express.Enabled**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True**，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True**，则将由应用程序确定其状态。但如果将该属性设置为 **False**，则强行禁用该控件。

#### **CommandBarButton.FaceId**

获取或设置某个 **CommandBarButton** 控件的图符的 ID 号。可读写。

**语法**

**express.FaceId**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

**FaceId** 属性确定一个命令栏按钮的外观而非其功能。**CommandBarControl** 对象的 **Id** 属性确定按钮的功能。

对于具有自定义图标的命令栏按钮，其 **FaceId** 属性值为 0（零）。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向自定义命令栏中添加一个命令栏按钮。虽然此按钮的图标与内置“制作图表”按钮的图标相同，但单击此按钮相当于单击“文件”菜单上的“打开”命令，这是因为此按钮的 ID 号为 23。 function test() {     let newBar = Application.CommandBars.Add("Custom2", Application.Enum.msoBarTop, null, true)     newBar.Visible = true      let con = newBar.Controls.Add(Application.Enum.msoControlButton, 23)     con.FaceId = 17 }` |

#### **CommandBarButton.Height**

获取或设置命令栏控件的高度。可读写。

**语法**

**express.Height**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置该控件宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。 function test() {     let myBar = Application.CommandBars.Item("Custom")     let barHeight = myBar.Height + 1     let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("Save").Id, null, null, true)     myControl.Height = barHeight * 2     myControl.Width = 50     myBar.Visible = true }` |

#### **CommandBarButton.HelpContextId**

获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.HelpFile**

获取或设置附加于 **CommandBarButton** 控件的“帮助”主题的文件名。可读写

**语法**

**express.HelpFile**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.HyperlinkType**

设置或获取一个 **MsoCommandBarButtonHyperlinkType** 常量，代表与指定的命令栏按钮相关联的超链接的类型。可读写。

**语法**

**express.HyperlinkType**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例检查名为“Custom”的命令栏上的指定命令栏按钮的 HyperlinkType 属性。如果 HyperlinkType 设置为 msoCommandBarButtonHyperlinkNone，则本示例会将该属性设置为 msoCommandBarButtonHyperlinkOpen，并将 URL 设置为 www.wps.cn。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     let myButton = myBar.Controls.Add(Application.Enum.msoControlButton)     myButton.FaceId = 277     myButton.HyperlinkType = Application.Enum.msoCommandBarButtonHyperlinkNone     if ( myButton.HyperlinkType > Application.Enum.msoCommandBarButtonHyperlinkOpen )     {         myButton.HyperlinkType = Application.Enum.msoCommandBarButtonHyperlinkOpen         myButton.TooltipText = "www.wps.cn"     } }` |

#### **CommandBarButton.Id**

获取内置的 **CommandBarButton** 控件的 ID。只读。

**语法**

**express.Id**

*express*   一个代表 **CommandBarButton** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 在本示例中，如果名为“Custom2”的命令栏上第一个控件的按钮的 ID 值小于 25，则更改其按钮图符。 */ function test() { 	let ctrl = Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1) 	if ( ctrl.Id < 25 ) {     	ctrl.FaceId = 17     	ctrl.Tag = "Changed control" 	} }  /*下面的示例将名为“Standard”的工具栏上每个控件的标题更改为该控件当前的 Id 属性值。*/  function test() { 	let ctl = Application.ActiveDocument.CommandBars.Item("Standard").Controls 	for (let i = 1; i <= ctl.Count; i++) {     	ctl.Item(i).Caption = String(ctl.Id) 	} }` |

#### **CommandBarButton.Index**

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarButton** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

第一个命令栏控件的位置为 1。**CommandBarControls** 集合中的分隔符不计算在内。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为“Custom2”的命令栏中搜索 ID 值为 23 的控件。如果能找到这样的控件，并且其索引号大于 5，则将该控件放置在命令栏上的第一个位置上。 function test() {     let myBar = Application.CommandBars.Item("Custom2")     let ctrl1 = myBar.FindControl(null,  23)      if( ctrl1.Index >  5 )     {         ctrl1.Move(null, 1)     } }` |

#### **CommandBarButton.IsPriorityDropped**

如果当前已根据使用统计信息和布局空间将 **CommandBarButton** 控件从菜单或工具栏中删除，则获取 **True**。（请注意，这与通过 Visible 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

如果 **IsPriorityDropped** 为 **True**，则 **Visible** 属性设置为 **True** 的控件将不能在个性化菜单或工具栏中立即显示出来。

为了确定何时将某一特定菜单项的 **IsPriorityDropped** 属性设置为 **True**，WPS Office会记录下该菜单项的使用次数的总和，并记录下用户使用了同一菜单中的其他菜单项而没有使用该特定菜单项的应用程序会话数。当该会话数达到一定的阈值之后，便会减少使用次数总和。当总次数减为零时，便将 **IsPriorityDropped** 设置为 **True**。编程人员不能设置会话值和阈值，也不能设置 **IsPriorityDropped** 属性。但他们可以使用 **AdaptiveMenus** 属性来禁用应用程序中特定菜单的自适应菜单。

为了确定何时将某一特定工具栏控件的 **IsPriorityDropped** 属性设置为 **True**，Office 维护了一个列表，记录下了该工具栏中各个控件的最后使用顺序。工具栏会在空间允许的前提下尽可能多地显示控件，控件按使用时间由近向远排列，最近使用过的控件显示在最前面。**Priority** 设置为 1 的控件将始终显示，而且为了显示这样的控件，在必要时工具栏还将换行。编程人员可使用 **Priority** 属性来确保始终显示特定的工具栏控件，或重新定位工具栏以便使其有足够的空间显示所有的控件。

使用下表可以预测在菜单项的 **IsPriorityDropped** 属性被设置为 **True** 之前，它仍将出现在个性化菜单中的会话数。

| 命令栏控件的使用次数 | 应用程序的会话数 |
| -------------------- | ---------------- |
| 0, 1                 | 3                |
| 2                    | 6                |
| 3                    | 9                |
| 4, 5                 | 12               |
| 6- 8                 | 17               |
| 9-13                 | 23               |
| 14-24                | 29               |
| 25 及 25 以上        | 31               |

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Left**

设置或获取指定的 **CommandBarButton** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Mask**

获取或设置代表 **CommandBarButton** 对象的屏蔽图像的 **IPictureDisp** 对象。屏蔽图像可决定按钮图像的透明部分。可读写。

**语法**

**express.Mask**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

在创建要作为屏蔽图像使用的图像时，所有要透明的区域应该为白色，所有要显示的区域应该为黑色。

通常在设置了 **CommandBarButton** 对象的图片后设置屏蔽。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarButton.OLEUsage**

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarButton** 控件。可读写。

**语法**

**express.OLEUsage**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

本属性允许用户在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载应用程序的各个命令栏控件在该 Office 应用程序中的表示方法。如果客户端和服务器均提供命令栏，那么命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 **OLEMenuGroup** 属性控制。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向名为“Tools”的命令栏中添加一个新按钮并设置其 OLEUsage 属性。 function test() {     let myControl = Application.CommandBars.Item("Tools").Controls.Add(Application.Enum.msoControlButton, null, null, null, true)     myControl.OLEUsage = Application.Enum.msoControlOLEUsageNeither }` |

#### **CommandBarButton.OnAction**

获取或设置一个 JS函数的名称，该过程在用户单击或更改 **CommandBarButton** 控件的值时运行。可读写。

**语法**

**express.OnAction**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为“Custom”的命令栏中添加一个命令栏控件。COM 加载项“FinanceAddIn”将在每次单击该控件时运行。 function test() {     let myBar = Application.CommandBars.Item("Custom")     let myControl = myBar.Controls.Add(Application.Enum.msoControlButton)     myControl.FaceId = 2     myControl.OnAction = "!<FinanceAddIn>"     myBar.Visible = true }` |

#### **CommandBarButton.Parameter**

获取或设置一个应用程序可用于执行 **CommandBarButton** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

如果为内置控件设置一个指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，那么该参数可用来为 Visual Basic 程序传送信息，或保存关于该控件的信息（类似于另一个 Tag 属性值）。

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将一个新参数赋予控件并将焦点置于新按钮上。 function test() {     let myControl = Application.CommandBars.Item("Custom").Controls.Item(4)     myControl.Copy(null, 1)     myControl.Parameter = "2"     myControl.SetFocus() }` |

#### **CommandBarButton.Parent**

获取 **CommandBarButton** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CommandBarButton** 对象的变量。

#### **CommandBarButton.Picture**

获取或设置一个代表 **CommandBarButton** 对象的图像的 **IPictureDisp** 对象。可读写。

**语法**

**express.Picture**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

在更改按钮图像时，还希望使用 **Mask** 属性设置屏蔽图像。屏蔽图像决定按钮图像透明的部分。通常在设置了 **CommandBarButton** 对象的图片后设置屏蔽。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarButton.Priority**

获取或设置 CommandBarButton 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。可读写。

**语法**

**express.Priority**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

有效的优先级数为从 0 至 7 之间的数字，默认值为 3。优先级为 1 表示不能从工具栏中删除该控件。而优先级为其他值的控件将被忽略。

菜单项的命令栏控件不能使用 **Priority** 属性。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.ShortcutText**

获取或设置当控钮出现在菜单、子菜单或快捷菜单上时 **CommandBarButton** 控件旁边显示的快捷键文字。可读写。

**语法**

**express.ShortcutText**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

只能为包含 **OnAction** 宏的命令栏按钮设置此属性。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在消息框中显示 ET 工作表菜单栏上“打开”命令（“文件”菜单）的快捷键文字。 alert (CommandBars.Item("Worksheet Menu Bar").Controls.Item("File").Controls.Item("New...").ShortcutText)` |

#### **CommandBarButton.State**

获取或设置 CommandBarButton 控件的外观。可读写。

**语法**

**express.State**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

内置命令栏按钮的 **State** 属性是只读的。**Type** 属性的值可作为 **MsoButtonState** 枚举中的值之一。

| ![img]()注释                                           |
| ------------------------------------------------------ |
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建一个名为“Custom”的命令栏并添加两个按钮，然后将左边的按钮设置为 msoButtonUp，并将右边的按钮设置为 msoButtonDown。 function test() {     // Add new command bar.     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     // Add 2 buttons to new command bar.     myBar.Controls.Add(Application.Enum.msoControlButton)     myBar.Controls.Add(Application.Enum.msoControlButton)     myBar.Visible = true           // Paste Bold button face and set State of first button.     let myControl1 = myBar.Controls.Item(1)     let imgSource = Application.CommandBars.FindControl(Application.Enum.msoControlButton, 113)     myControl1.State = Application.Enum.msoButtonUp          // Paste italic button face and set State of second button.     let myControl2 = myBar.Controls.Item(2)     imgSource = Application.CommandBars.FindControl(Application.Enum.msoControlButton, 114)     myControl2.State = Application.Enum.msoButtonDown }` |

#### **CommandBarButton.Style**

获取或设置 **CommandBarButton** 控件的显示方式。可读写。

**语法**

**express.Style**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

**Style** 属性的值可作为 **MsoButtonStyle** 枚举中的值之一。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Tag**

获取或设置有关 **CommandBarButton** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

*express*   一个代表 **CommandBarButton** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例可实现的功能为：将自定义工具栏上的按钮的标记设置为“Spelling Button”，并在消息框中显示此标记。 */ function test() {     Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1).Tag = "Spelling Button" 	alert(Application.ActiveDocument.CommandBars.Item("Custom").Controls.Item(1).Tag) }   ` |

#### **CommandBarButton.TooltipText**

获取或设置 **CommandBarButton**的**“屏幕提示”**中显示的文本。可读写。

**语法**

**express.TooltipText**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

默认情况下，**Caption** 属性的值将用作**“屏幕提示”**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Top**

获取指定的 **CommandBarButton** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Type**

获取 **CommandBarButton** 控件的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarButton.Visible**

获取或设置 **CommandBarButton** 控件的 **Visible** 属性。如果 **CommandBarButton** 可见，则此属性为 **true**。可读写。

**语法**

**express.Visible**

*express*   一个代表 **CommandBarButton** 对象的变量。

#### **CommandBarButton.Width**

获取或设置指定的 **CommandBarButton** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

*express*   一个代表 **CommandBarButton** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

适用环境：web

适用平台：windows/linux