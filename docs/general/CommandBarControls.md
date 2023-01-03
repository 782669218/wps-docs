#### **CommandBarControls**



**CommandBarControl** 对象的集合，这些对象代表某个命令栏上的命令栏控件。

**说明**

使用 **Controls** 属性可返回 **CommandBarControls** 集合。下面的示例将名为“Standard”的工具栏上每个控件的标题更改为该控件当前的 **Id** 属性值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let cmdBars = Application.ActiveDocument.CommandBars 	for(let i = 1; i <= cmdBars.Item("Standard").Controls.Count; i++){    	 cmdBars.Item("Standard").Controls.Item(i).Caption = String(cmdBars.Item("Standard").Controls.Item(i).Id) 	} }` |

使用 **Add** 方法可在 **CommandBarControls** 集合中添加一个新的命令栏控件。本示例将在名为“Custom”的命令栏中添加一个新的空白按钮。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myBlankBtn = Application.ActiveDocument.CommandBars.Item("Custom").Controls.Add()` |

使用 Controls(index) 可返回一个 **CommandBarControl**、**CommandBarButton**、**CommandBarComboBox** 或 **CommandBarPopup** 对象，其中 *index* 是该控件的标题或索引号。下面的示例将名为“Standard”的命令栏中的第一个控件复制到名为“Custom”的命令栏中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let myCustomBar = CommandBars.Item("Custom") 	let myControl = CommandBars.Item("Standard").Controls.Item(1)         myControl.Copy(myCustomBar, 1) }` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 新建一个 **CommandBarControl** 对象并将其添加到指定命令栏上的控件集合中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 获取 **CommandBarControls** 集合中的一个 **CommandBarControl**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E9%80%9A%E7%94%A8%20API%20%E5%8F%82%E8%80%83/CommandBarControl/CommandBarControl%20.htm#jsObject_CommandBarControl)对象。只读。 |

**属性**

|                                                              | 名称      | 说明                           |
| ------------------------------------------------------------ | --------- | ------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count** | 获取命令栏上的控件数量。只读。 |

**成员方法**

#### **CommandBarControls.Add**

新建一个 **CommandBarControl** 对象并将其添加到指定命令栏上的控件集合中。

**语法**

**express.Add(Type, Id , Parameter, Before, Temporary)**

*express*   一个代表 **CommandBarControls** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*      | 可选          | **Variant**  | 要添加到指定命令栏中的控件类型。可以为以下 MsoControl 常量之一：msoControlButton、msoControlEdit、msoControlDropdown、msoControlComboBox 或 msoControlPopup。 |
| *Id*        | 可选          | **Variant**  | 指定内置控件的整数。如果该参数为 1，或者忽略该参数，将在命令栏中添加一个空的指定类型的自定义控件。 |
| *Parameter* | 可选          | **Variant**  | 对于内置控件，容器应用程序使用该参数运行命令。对于自定义控件，可以使用该参数向 Visual Basic 过程传递信息，也可以用它存储有关控件的信息（类似于第二个 Tag 属性值）。 |
| *Before*    | 可选          | **Variant**  | 一个指示新控件在命令栏上位置的数值。新控件将插入到位于此位置的控件之前。如果忽略该参数，控件将添加到指定命令栏的末端。 |
| *Temporary* | 可选          | **Variant**  | 设置为 true 将使新控件为临时控件。临时控件在关闭容器应用程序时自动删除。默认值为 false。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建包含剪切、复制和粘贴按钮（控件）的自定义编辑工具栏。 function test() { 	let customBar = Application.ActiveDocument.CommandBars.Add("Custom") 	let newButton = customBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Edit").Controls.Item("Cut").Id) 	let newButton1 = customBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Edit").Controls.Item("Copy").Id) 	let newButton2 = customBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Edit").Controls.Item("Paste").Id)     customBar.Visible = true }` |

#### **CommandBarControls.Item**

获取 **CommandBarControls** 集合中的一个 **CommandBarControl**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E9%80%9A%E7%94%A8%20API%20%E5%8F%82%E8%80%83/CommandBarControl/CommandBarControl%20.htm#jsObject_CommandBarControl)对象。只读。

**语法**

**express.Item(Index)**

*express*   一个代表 **CommandBarControls** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Index*  | 必选          | **Variant**  | 要返回的对象的名称或索引号。 |

**成员属性**

#### **CommandBarControls.Count**

获取命令栏上的控件数量。只读。

**语法**

**express.Count**

*express*   一个代表 **CommandBarControls** 对象的变量。

适用环境：web

适用平台：windows/linux