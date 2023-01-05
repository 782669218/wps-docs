**Application**



代表整个 WPP 应用程序。 

**说明**

**Application** 对象包括：

- 应用程序范围内的设置和选项（例如，当前打印机的名称）。
- 用于返回顶层对象的属性，例如 **ActivePresentation** 和 **Windows**。

在WPS宏编辑器中编写要在WPP 中运行的代码时，**Application** 对象的下列属性可以在没有对象限定符的情况下使用：**ActivePresentation**、**ActiveWindow**、**AddIns**、**Presentations**、**SlideShowWindows** 和 **Windows**。例如，可以编写ActiveWindow.Height = 200 来代替Application.ActiveWindow.Height = 200。（仅WPS宏编辑器）

**方法**

|                                                              | 名称         | 说明                                                |
| ------------------------------------------------------------ | ------------ | --------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate** | 激活指定的对象。                                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Quit**     | 退出 WPP。该方法相当于单击“文件”菜单上的“退出WPP”。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Run**      | 运行JS宏。                                          |

**属性**

|                                                              | 名称                                | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**                          | 返回指定的窗格或窗口是否处于激活的状态。只读。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveEncryptionSession**         | 此对象或成员加入对象模型中的时间太晚，以至于还没有完整的文档说明。 只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePresentation**              | 返回一个 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象，该对象代表当前窗口中打开的演示文稿。只读。请注意，如果嵌入的演示文稿处于活动状态，则 **ActivePresentation** 属性将返回嵌入的演示文稿。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePrinter**                   | 返回当前打印机的名称。只读。**String** 类型。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWindow**                    | 返回一个代表活动文档窗口的 **DocumentWindow** 对象。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Assistance**                      | 获取 WPS Office IAssistance 对象的引用，该对象为开发人员提供了一种可以为 WPS Office 用户创建自定义帮助体验的方法。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoCorrect**                     | 返回一个代表 WPP 中的“自动更正”功能的 **AutoCorrect** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Build**                           | 返回WPP 内部版本号。只读 **String** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**                         | 返回在应用程序窗口的标题栏中显示的文本。**String** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **COMAddIns**                       | 返回对 WPP 中当前加载的组件对象模型 (COM) 加载宏的引用。在“WPP选项”对话框的**“加载宏”**选项卡中列出了这些加载宏。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBars**                     | 返回一个代表WPP 中所有命令栏的 **CommandBars** 集合。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                         | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultWebOptions**               | 返回一个 **DefaultWebOptions** 对象，该对象包含将部分或全部演示文稿作为网页发布或者打开一个网页时 WPP 使用的全局应用程序级属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAlerts**                   | 设置或返回 **PpAlertLevel** 常量，该常量代表 WPP 在运行宏时是否显示警告。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayDocumentInformationPanel** | 返回或设置是否在 WPP 用户界面中显示“文档属性”面板。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayGridLines**                | **MsoTrue** 属性值用于显示 WPP 中的网格线。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FeatureInstall**                  | 返回或设置 WPP 如何处理对所需功能尚未安装的方法和属性的调用。可读/写 **MsoFeatureInstall** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FileDialog**                      | 返回一个 **FileDialog** 对象，该对象表示文件对话框的单个实例。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                          | 以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageSettings**                | 返回一个 **LanguageSettings** 对象，该对象包含 WPP 中有关语言设置的信息。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                            | 返回或设置一个 **Single** 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                            | 返回应用程序的名称字符串。**String** 类型，只读 。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NewPresentation**                 | 返回一个 **NewFile** 对象，该对象代表“新建演示文稿”任务窗格中列出的演示文稿。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OperatingSystem**                 | 返回操作系统名称。只读 **String** 类型。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Options**                         | 返回一个 **Options** 对象，该对象代表 WPP 中的应用程序选项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**                            | 返回 **String**，该值代表到指定的 **Application** 对象的路径，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Presentations**                   | 返回一个 **Presentations** 集合，该集合代表所有打开的演示文稿。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **Presentation**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProductCode**                     | 返回 WPP 的全局唯一标识符 (GUID)。用户可能会用到 GUID，例如，当使用程序调用应用程序编程接口 (API) 时。只读 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowStartupDialog**               | 属性值为 **MsoTrue** 时，启动 WPP 后显示“新建演示文稿”任务窗格。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowWindowsInTaskbar**            | 决定是否每个打开的演示文稿都有单独的 Windows 任务栏按钮。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowWindows**                | 返回一个 **SlideShowWindows** 集合，该集合代表所有打开的幻灯片放映窗口。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **SlideShowWindow**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                             | 返回或设置一个 **Single** 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Version**                         | 返回WPP 版本号。只读 **String** 类型。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                         | 返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                           | 以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**                         | 返回一个代表所有打开的文档窗口的 [**DocumentWindow**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/DocumentWindows/DocumentWindows%20.htm#jsObject_DocumentWindows) 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**                     | 返回或设置指定窗口的状态。可读/写。**PpWindowState** 类型。  |

**成员方法**

#### **Application.Activate**

激活指定的对象。

**语法**

**express.Activate()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                             |
| -------------------------------------------------------- |
| `/*本示例会激活WPP程序当前窗口*/ Application.Activate()` |

#### **Application.Quit**

退出 WPP。该方法相当于单击“文件”菜单上的“退出WPP”。

**语法**

**express.Quit()**

*express*   一个代表 **Application** 对象的变量。

**说明**

若要避免提示保存更改，可在调用 **Quit** 方法之前使用 **Save** 或 **SaveAs** 方法保存所有打开的演示文稿。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例保存所有打开的演示文稿，然后退出WPP。*/ function test() {     let app = Application     for(let i=1; i <= app.Presentations.Count; i++)      {         app.Presentations.Item(i).Save()     }     app.Quit() }` |

#### **Application.Run**

运行JS宏。

**语法**

**express.Run(MacroName, safeArrayOfParams)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *MacroName*         | 必选          | **String**   | 要运行的宏的名称。该字符串可包含下列内容：后跟感叹号 (!) 的加载演示文稿或加载宏文件名、后跟句点 (.) 的有效模块名称以及宏名称。例如，“MyPres.ppt!Module1.Test”是一个有效的 MacroName 值。 |
| *safeArrayOfParams* | 必选          | **Variant**  | 要传递给该过程的参数。不能为此参数指定一个对象，也不能对此方法使用命名参数。参数必须按位置传递。 |

**说明**

宏可能包含病毒，因此在运行宏时要格外小心。请采用下列预防措施：在计算机上运行最新的防病毒软件；将宏安全级别设置为“高”；清除“信任所有安装的加载项和模板”复选框；使用数字签名；维护受信任的发布者的列表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在本示例中，Main 过程定义一个数组，然后运行宏 TestPass，将该数组作为参数传递。*/ function Main() {     let x = []     x[1] = "hi"     x[2] = 7     Application.Run("TestPass", x) }  function TestPass(x) {     alert(x[1])     alert(x[2]) }` |

**成员属性**

#### **Application.Active**

返回指定的窗格或窗口是否处于激活的状态。只读。**MsoTriState** 类型。

**语法**

**express.Active**

*express*   一个代表 **Application** 对象的变量。

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
| `//以下示例检查演示文稿文件 "test.ppt" 是否位于当前窗口。如果不是，则将当前处于激活状态的演示文稿的名称保存在变量 oldWin 中，并激活 "test.ppt" 演示文稿。 function test(){     let window = Application.Presentations.Item("test.ppt").Windows.Item(1)     if(!window.Active) {         let oldWin = Application.ActiveWindow         window.Activate()     } }` |

#### **Application.ActiveEncryptionSession**

此对象或成员加入对象模型中的时间太晚，以至于还没有完整的文档说明。 只读。

**语法**

**express.ActiveEncryptionSession**

*express*   一个代表 **Application** 对象的变量。

#### **Application.ActivePresentation**

返回一个 [Presentation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentation/Presentation%20.htm#jsObject_Presentation)对象，该对象代表当前窗口中打开的演示文稿。只读。请注意，如果嵌入的演示文稿处于活动状态，则 **ActivePresentation** 属性将返回嵌入的演示文稿。

**语法**

**express.ActivePresentation**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将加载的演示文稿保存到D盘根目录下名为“TestFile”的文件中。*/ function test() {     let MyPath = "D:\TestFile"     Application.ActivePresentation.SaveAs(MyPath) } ` |

#### **Application.ActivePrinter**

返回当前打印机的名称。只读。**String** 类型。

**语法**

**express.ActivePrinter**

*express*   一个代表 **Application** 对象的变量。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例显示当前打印机的名称。 alert("The name of the active printer is " + Application.ActivePrinter)` |

#### **Application.ActiveWindow**

返回一个代表活动文档窗口的 **DocumentWindow** 对象。只读。

**语法**

**express.ActiveWindow**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口最小化。*/ Application.ActiveWindow.WindowState = ppWindowMinimized` |

#### **Application.Assistance**

获取 WPS Office IAssistance 对象的引用，该对象为开发人员提供了一种可以为 WPS Office 用户创建自定义帮助体验的方法。只读。

**语法**

**express.Assistance**

*express*   一个代表 **Application** 对象的变量。

#### **Application.AutoCorrect**

返回一个代表 WPP 中的“自动更正”功能的 **AutoCorrect** 对象。

**语法**

**express.AutoCorrect**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例禁止显示“自动更正选项”和“自动版式选项”按钮。 function HideAutoCorrectOpButton() {     let ac = Application.AutoCorrect     ac.DisplayAutoCorrectOptions = msoFalse     ac.DisplayAutoLayoutOptions = msoFalse }` |

#### **Application.Build**

返回WPP 内部版本号。只读 **String** 类型。

**语法**

**express.Build**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                               |
| ---------------------------------------------------------- |
| `/*本示例显示WPP 内部版本号。*/ alert(Application.Build) ` |

#### **Application.Caption**

返回在应用程序窗口的标题栏中显示的文本。**String** 类型，可读/写。

**语法**

**express.Caption**

*express*   一个代表 **Application** 对象的变量。

#### **Application.COMAddIns**

返回对 WPP 中当前加载的组件对象模型 (COM) 加载宏的引用。在“WPP选项”对话框的**“加载宏”**选项卡中列出了这些加载宏。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.COMAddIns**

*express*   一个代表 **Application** 对象的变量。

#### **Application.CommandBars**

返回一个代表WPP 中所有命令栏的 **CommandBars** 集合。只读。

**语法**

**express.CommandBars**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **Application** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例显示关于 myObject 创建者的消息。 function test(){     let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)     if(myObject.Creator == parseInt(50575054,16)) {        alert("This is a WPP object")     }     else {     alert("This is not a WPP object")     } }` |

#### **Application.DefaultWebOptions**

返回一个 **DefaultWebOptions** 对象，该对象包含将部分或全部演示文稿作为网页发布或者打开一个网页时 WPP 使用的全局应用程序级属性。只读。

**语法**

**express.DefaultWebOptions**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例检查默认的文档编码是否为 Western。如果是 Western，则字符串 strDocEncoding 将根据它进行设置。 function test(){     let objAppWebOptions = Application.DefaultWebOptions     if(objAppWebOptions.Encoding == msoEncodingWestern) {         let strDocEncoding = "Western"     } }` |

#### **Application.DisplayAlerts**

设置或返回 **PpAlertLevel** 常量，该常量代表 WPP 在运行宏时是否显示警告。可读/写。

**语法**

**express.DisplayAlerts**

*express*   一个代表 **Application** 对象的变量。

**说明**

宏停止运行后，**DisplayAlerts** 属性的值不会重新设置；该值将在会话过程中保持不变。在会话过程之间不会保存该值，所以在WPP 开始运行时，该值重新设置为 **ppAlertsNone**。

| **PpAlertLevel** 可以是下列 **PpAlertLevel** 常量之一。      |
| ------------------------------------------------------------ |
| **ppAlertsAll** 显示所有的消息框和警告；将错误返回宏。       |
| **ppAlertsNone** 默认值。不显示任何警告或消息框。如果运行宏时遇到消息框，则选中默认值并继续运行宏。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码行指示WPP 显示所有的消息框和警告，并将错误返回宏。 Application.DisplayAlerts = ppAlertsAll` |

#### **Application.DisplayDocumentInformationPanel**

返回或设置是否在 WPP 用户界面中显示“文档属性”面板。可读/写。

**语法**

**express.DisplayDocumentInformationPanel**

*express*   一个代表 **Application** 对象的变量。

#### **Application.DisplayGridLines**

**MsoTrue** 属性值用于显示 WPP 中的网格线。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayGridLines**

*express*   一个代表 **Application** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例切换WPP 中网格线的显示状态。 function ToggleGridLines() {     let a = Application     if(a.DisplayGridLines == msoTrue) {         a.DisplayGridLines = msoFalse 	}     else {         a.DisplayGridLines = msoTrue     } }` |

#### **Application.FeatureInstall**

返回或设置 WPP 如何处理对所需功能尚未安装的方法和属性的调用。可读/写 **MsoFeatureInstall** 类型。

**语法**

**express.FeatureInstall**

*express*   一个代表 **Application** 对象的变量。

**说明**

可以使用 **msoFeatureInstallOnDemandWithUI** 常量防止用户误以为在安装某功能时系统没有反应。使用带有错误捕获例程的 **msoFeatureInstallNone** 常量可防止安装最终用户功能。

| ![img]() 注释                                                |
| ------------------------------------------------------------ |
| 如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **FeatureInstall** 属性设置是什么，该模板都不会自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须首先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的“添加/删除程序”图标）安装WPP 的附加设计模板。 |

| MsoFeatureInstall 可以是下列 MsoFeatureInstall 常量之一。    |
| ------------------------------------------------------------ |
| **msoFeatureInstallNone** 默认值。调用未安装的功能时会产生可捕获的运行时自动错误。 |
| **msoFeatureInstallOnDemand** 显示对话框提示用户安装新功能。 |
| **msoFeatureInstallOnDemandWithUI** 安装时显示进度条。不提示用户安装新功能。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例检查 FeatureInstall 属性的值。如果该属性设置为 msoFeatureInstallNone，则代码显示一个消息框，以询问用户是否要更改属性设置。如果用户回答“是”，则此属性将设置为 msoFeatureInstallOnDemand。 function test(){     let a = Application     if(a.FeatureInstall == msoFeatureInstallNone) {         let Reply = alert("Uninstalled features for this application " + "\r\n"         + "may cause a run-time error when called." + "\r\n" + "\r\n" + "Would you like to change this setting"          + "\r\n"  + "to automatically install missing features when called?" , 52, "Feature Install Setting")         if(Reply == 6) {             a.FeatureInstall = msoFeatureInstallOnDemand         }     } }` |

#### **Application.FileDialog**

返回一个 **FileDialog** 对象，该对象表示文件对话框的单个实例。

**语法**

**express.FileDialog**

*express*   一个代表 **Application** 对象的变量。

**说明**

| **MsoFileDialogType** 可以是下列 **MsoFileDialogType** 常量之一。 |
| ------------------------------------------------------------ |
| **msoFileDialogFilePicker**                                  |
| **msoFileDialogFolderPicker**                                |
| **msoFileDialogOpen**                                        |
| **msoFileDialogSaveAs**                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例显示“另存为”对话框。 function ShowSaveAsDialog() {     let dlgSaveAs = Application.FileDialog( msoFileDialogSaveAs)     dlgSaveAs.Show() }` |

#### **Application.Height**

以磅为单位返回或设置指定对象的高度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。

**语法**

**express.Height**

*express*   一个代表 **Application** 对象的变量。

**说明**

一个 **Height**Shape 对象的 **Shape**Height 属性返回或设置指定形状向前的表面的高度。该尺寸不包括阴影或三维效果。

#### **Application.LanguageSettings**

返回一个 **LanguageSettings** 对象，该对象包含 WPP 中有关语言设置的信息。只读。

**语法**

**express.LanguageSettings**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Left**

返回或设置一个 **Single** 类型值，该值代表从文档、应用程序和幻灯片放映窗口的左边缘到应用程序窗口用户区域左边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Left**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Name**

返回应用程序的名称字符串。**String** 类型，只读 。

**语法**

**express.Name**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则对形状集合对象Shapes调用Item("Rectangle 2")将返回对该形状的引用。详情见示例。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let name = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).Name;     let shape = Application.ActivePresentation.Slides.Item(1).Shapes.Item(name)     alert(shape.Name) }` |

#### **Application.NewPresentation**

返回一个 **NewFile** 对象，该对象代表“新建演示文稿”任务窗格中列出的演示文稿。只读。

**语法**

**express.NewPresentation**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例会在窗格中最后部分的底部列出“新建演示文稿”任务窗格上的演示文稿。  function CreateNewPresentationListItem() {     Application.NewPresentation.Add("C:\\Presentation.ppt")     Application.CommandBars.Item(1).Visible = true }` |

#### **Application.OperatingSystem**

返回操作系统名称。只读 **String** 类型。

**语法**

**express.OperatingSystem**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例检测 OperatingSystem 属性以确定WPP 是否运行在 32 位版本的 Windows 中。 function test(){     let os = Application.OperatingSystem     if((os.indexOf("Windows (32-bit)") != -1) != 0) {         alert("Running a 32-bit version of  Windows")     } }` |

#### **Application.Options**

返回一个 **Options** 对象，该对象代表 WPP 中的应用程序选项。

**语法**

**express.Options**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//使用 Options 属性可返回 Options 对象。以下示例设置WPP 的三个应用程序选项。 function TogglePasteOptionsButton() {     let options = Application.Options     if(options.DisplayPasteOptions == false) {         options.DisplayPasteOptions = true     } }` |

#### **Application.Path**

返回 **String**，该值代表到指定的 **Application** 对象的路径，只读。

**语法**

**express.Path**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `/*本示例显示WPP程序的路径。*/ alert(Application.Path)` |

#### **Application.Presentations**

返回一个 **Presentations** 集合，该集合代表所有打开的演示文稿。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **Presentation**。

**语法**

**express.Presentations**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开名为“TEST”的演示文稿。*/ Application.Presentations.Open("C:\\Users\\wps\\Desktop\\TEST.pptm")  /*本示例将第一个演示文稿以“Year-End Report.ppt”名称保存在D盘。*/ Application.Presentations.Item(1).SaveAs("D:/Year-End Report.pptx")  /*本示例关闭演示文稿“Year-end report”。*/ Application.Presentations.Item("Year-End Report.ppt").Close()` |

#### **Application.ProductCode**

返回 WPP 的全局唯一标识符 (GUID)。用户可能会用到 GUID，例如，当使用程序调用应用程序编程接口 (API) 时。只读 **String** 类型。

**语法**

**express.ProductCode**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `/*本示例将显示GUID。*/ alert(Application.ProductCode)` |

#### **Application.ShowStartupDialog**

属性值为 **MsoTrue** 时，启动 WPP 后显示“新建演示文稿”任务窗格。可读/写。

**语法**

**express.ShowStartupDialog**

*express*   一个代表 **Application** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue** 不适用于此属性。                         |
| **msoFalse** 隐藏“新建演示文稿”侧窗格。               |
| **msoTriStateMixed** 不适用于此属性。                 |
| **msoTriStateToggle** 不适用于此属性。                |
| **msoTrue** 默认值。显示“新建演示文稿”侧窗格。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码行在WPP 启动时禁用“新建演示文稿”任务窗格。 Application.ShowStartupDialog = msoFalse` |

#### **Application.ShowWindowsInTaskbar**

决定是否每个打开的演示文稿都有单独的 Windows 任务栏按钮。可读/写 **MsoTriState** 类型。

**语法**

**express.ShowWindowsInTaskbar**

*express*   一个代表 **Application** 对象的变量。

**说明**

设置为 **True** 时，此属性将模拟单文档界面 (SDI) 的外观，这样便于在打开的演示文稿之间切换。但是，如果在使用多个演示文稿的同时还打开了其他应用程序，则可能希望将此属性设置为 **False**，以避免任务栏上布满多余的按钮。

仅当 WPS Office与 Windows Update 或 Windows 系统一同使用时此属性可用。

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**                                                 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 默认值。每个打开的演示文稿都有单独的 Windows 任务栏按钮。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例指定每个打开的演示文稿都不具有单独的 Windows 任务栏按钮。 Application.ShowWindowsInTaskbar = msoFalse` |

#### **Application.SlideShowWindows**

返回一个 **SlideShowWindows** 集合，该集合代表所有打开的幻灯片放映窗口。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象 **SlideShowWindow**。

**语法**

**express.SlideShowWindows**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个窗口中运行幻灯片放映，并设置该幻灯片放映窗口的高度和宽度。*/ function test() {     let a = Application     a.Presentations.Item(1).SlideShowSettings.Run()     let windows = a.SlideShowWindows.Item(1)     windows.Height = 250     windows.Width = 250 }` |

#### **Application.Top**

返回或设置一个 **Single** 类型值，该值代表文档、应用程序和幻灯片放映窗口的上边缘到应用程序窗口用户区域上边缘的距离（以磅为单位）。将该属性设为一个很大的正值或负值会将整个窗口移出桌面。可读/写。

**语法**

**express.Top**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Version**

返回WPP 版本号。只读 **String** 类型。

**语法**

**express.Version**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个消息框中显示WPP 版本号和内部版本号以及操作系统名称。*/ function test() {     let a = Application     alert("Welcome to WPP version " + a.Version + ", build " + a.Build + ", running on " + a.OperatingSystem + "!") }` |

#### **Application.Visible**

返回或设置指定对象的可见性或应用于指定对象的格式。可读/写。**MsoTriState** 类型。

**语法**

**express.Visible**

*express*   一个代表 **Application** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**：指定对象或对象格式是可见的。     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例为当前演示文稿中第一张幻灯片的第三个形状设置阴影的水平偏移量和垂直偏移量，使其向右偏移 5 磅并向上偏移 3 磅。如果该形状原来没有阴影，本例将会为其添上阴影。 function test(){     let myDocument = Application.ActivePresentation.Slides.Item(1)     let shadow = myDocument.Shapes.Item(3).Shadow     shadow.Visible = msoTrue     shadow.OffsetX = 5     shadow.OffsetY = -3 }` |

#### **Application.Width**

以磅为单位返回或设置指定对象的宽度。用于 **Master** 对象时只读，**Single** 类型；用于所有其他对象时可读/写，**Single** 类型。

**语法**

**express.Width**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Windows**

返回一个代表所有打开的文档窗口的 [**DocumentWindow**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/DocumentWindows/DocumentWindows%20.htm#jsObject_DocumentWindows) 集合。只读。

**语法**

**express.Windows**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test() {     let aw = Application.Windows     for(let i = aw.Count; i >= 2; i--)      {         aw.Item(i).Close()     } }` |

#### **Application.WindowState**

返回或设置指定窗口的状态。可读/写。**PpWindowState** 类型。

**语法**

**express.WindowState**

*express*   一个代表 **Application** 对象的变量。

**说明**

| PpWindowState 可以是下列 PpWindowState 常量之一。 |
| ------------------------------------------------- |
| **ppWindowMaximized**                             |
| **ppWindowMinimized**                             |
| **ppWindowNormal**                                |

当窗口状态为 **ppWindowNormal** 时，该窗口既未最大化也未最小化。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前窗口最小化。*/ Application.ActiveWindow.WindowState = ppWindowMinimized` |

适用环境：web

适用平台：windows/linux