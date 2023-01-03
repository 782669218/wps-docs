**AddIn**



代表已加载或未加载的单个加载项。

**说明**

**AddIn** 对象是 **AddIns** 集合的成员。**AddIns** 集合包含与 WPP 相关的所有可用加载项，无论这些加载项是否加载。该集合不包含组件对象模型 (COM) 加载项 （COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。）。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 Application 对象，该对象表示指定对象的创建者。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoLoad**    | 决定指定的加载宏是否在每次启动WPP 时自动加载。MsoTriState 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**    | 返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。String 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Loaded**      | 确定是否已加载指定加载项。MsoTriState 类型，可读/写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 已注册的文件类型的加载项名称（标题）。String 类型，只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回 String，该值代表到指定的 AddIn、Application 或 Presentation 对象的路径，或者后跟 MotionEffect 对象的路径。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Registered**  | 返回指定的加载宏是否注册到 Windows 注册表中。MsoTriState 类型，可读/写。 |

**成员属性**

#### **AddIn.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path & "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){ 　　　　for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {    　　　　 if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {     　　　　    alert(ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)   　　　　  } 　　　　} }` |

#### **AddIn.AutoLoad**

决定指定的加载宏是否在每次启动WPP 时自动加载。MsoTriState 类型，可读/写。

**语法**

**express.AutoLoad**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

如果将该属性设置为 **msoTrue**，则自动将 **Registered** 属性设置为 **msoTrue**。

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| ---------------------------------------------- |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 每次启动WPP 时，自动加载指定的加载宏。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示每次启动WPP 时自动加载的每个加载宏的名称。*/ function test(){     let afound     for(let i=1;i <= Application.AddIns.Count;i++) {  　　　　   if(Application.AddIns.Item(i).AutoLoad) {   　　　　      alert(Application.AddIns.Item(i).Name)    　　　　     afound = true    　　　　 } 　　　　} 　　　　if(afound != true) { 　　　　    alert("No add-ins were loaded automatically.") 　　　　} }` |

本示例指定名为“MyTools”的加载宏在每次启动WPP 时自动加载。

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `Application.AddIns.Item("mytools").AutoLoad = msoTrue` |

#### **AddIn.FullName**

返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。String 类型。

**语法**

**express.FullName**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

此属性与 Path 属性等效，附加上当前文件系统分隔符和 Name 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示每一个可用加载宏的路径及文件名。*/ function test(){ 　　　　for(let i=1;i <= Application.AddIns.Count;i++) {    　　　　 alert(Application.AddIns.Item(i).FullName) 　　　　} }` |

以下示例显示当前演示文稿的路径和文件名（假设演示文稿已保存）。

| 示例代码复制                                     |
| ------------------------------------------------ |
| `alert(Application.ActivePresentation.FullName)` |

#### **AddIn.Loaded**

确定是否已加载指定加载项。MsoTriState 类型，可读/写。

**语法**

**express.Loaded**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| ---------------------------------------------- |
| msoCTrue                                       |
| msoFalse                                       |
| msoTriStateMixed                               |
| msoTriStateToggle                              |
| msoTrue 指定加载项已加载。在“加载项”选项卡中。 |

**示例**

以下示例将 MyTools.ppa 添加到**“加载项”**选项卡中的列表，然后加载它。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.AddIns.Add("c:\\my documents\\mytools.ppa").Loaded = msoTrue` |

以下示例卸载名为“MyTools”的加载项。

| 示例代码复制                                           |
| ------------------------------------------------------ |
| `Application.AddIns.Item("mytools").Loaded = msoFalse` |

#### **AddIn.Name**

已注册的文件类型的加载项名称（标题）。String 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

如果包含某对象的集合的 Item 方法采用 Variant 类型的参数，则可以将该对象的名称与 Item 方法联合使用来返回对该对象的引用。例如，如果某个形状的 Name 属性值为“Rectangle 2”，则 .Shapes.Item("Rectangle 2") 将返回对该形状的引用。

#### **AddIn.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){     　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     　　　　let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     　　　　addShp.TextRange.Text = "Test text"     　　　　addShp.Parent.Rotation = 45     }` |

#### **AddIn.Path**

返回 String，该值代表到指定的 AddIn、Application 或 Presentation 对象的路径，或者后跟 MotionEffect 对象的路径。只读。

**语法**

**express.Path**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

如果使用此属性返回一个尚未保存的演示文稿的路径，则会得到一个空字符串。如果使用此属性返回未加载的加载宏的路径，则会导致错误。
路径中不包括最后的反斜线 (\) 或指定对象的名称。使用 Presentation 对象的 Name 属性可返回不带路径的文件名，使用 FullName 属性则可返回带路径的文件名。
MotionEffect 对象返回的 String 类型值是一个特殊路径，它使用与 VML 路径说明相同的语法描述移动效果在 From 和 To 之间沿行的路径。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前演示文稿保存到与WPP 相同的文件夹中。 */ function test(){ 　　　　let fName = Application.Path + "\\test presentation" 　　　　Application.ActivePresentation.SaveAs(fName) }` |

#### **AddIn.Registered**

返回指定的加载宏是否注册到 Windows 注册表中。MsoTriState 类型，可读/写。

**语法**

**express.Registered**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。   |
| ----------------------------------------------- |
| msoCTrue                                        |
| msoFalse                                        |
| msoTriStateMixed                                |
| msoTriStateToggle                               |
| msoTrue 将指定的加载宏注册到 Windows 注册表中。 |

**示例**

本示例将名为“MyTools”的加载宏注册到 Windows 注册表中。

| 示例代码复制                                              |
| --------------------------------------------------------- |
| `Application.AddIns.Item("MyTools").Registered = msoTrue` |

适用环境：web

适用平台：windows/linux