**AddIn** 对象

代表已经安装或尚未安装的单个加载项。**AddIn** 对象是 **AddIns** 集合的成员。**AddIns** 集合包含可供 WPS 使用的所有加载项，不论它们当前是否加载。**AddIns** 集合中包括在**“工具”**菜单的**“模板和加载项”**对话框中显示的共用模板或 WPS 加载项库 (WLL)。

**说明**

可以使用 **AddIns**( *index*) 返回单个 **AddIn** 对象（其中 *index* 是加载项的名称或索引编号）。名称的拼写必须和**“模板和加载项”**对话框中显示的完全一样（但大小写可以不一样）。以下示例将 Letter.dot 模板作为共用模板加载。


| 示例代码复制                                             |
| -------------------------------------------------------- |
| ```Application.AddIns.Item("Letter.dot").Installed = true``` |

索引编号代表加载项在**“模板和加载项”**对话框中的加载项列表上的位置。下列指令显示第一个可用加载项的路径。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```function test() {     if (Application.AddIns.Count >= 1) {         alert(Application.AddIns.Item(1).Path)     } }``` |

以下示例在活动文档开始处创建一个加载项列表。该列表包含了每一个有效加载项的名称、路径和加载状态。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```function test() {     let rng = Application.ActiveDocument.Range(0, 0)     rng.InsertAfter("Name" + "\t" + Path + "\t" + Installed)     rng.InsertParagraphAfter()     for (let i = 1; i <= Application.AddIns.Count; i++) {         rng.InsertAfter(Application.AddIns.Item(i).Name + "\t" + Application.AddIns.Item(i).Path + "\t" + Application.AddIns.Item(i).Installed)         rng.InsertParagraphAfter()     }     rng.ConvertToTable() }``` |

使用 **Add** 方法将一个加载项添加到可用加载项列表上并且使用 *Install* 参数安装它（可选）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```Application.AddIns.Add("C:\\Templates\\Other\\Letter.dot", true)``` |

要安装可用加载项列表中显示的加载项，请使用 **Installed** 属性。

| 示例代码复制                                             |
| -------------------------------------------------------- |
| ```Application.AddIns.Item("Letter.dot").Installed = true``` |

| 注释                                                         |
| ------------------------------------------------------------ |
| 使用 **Compiled** 属性判断一个 **AddIn** 对象是模板还是 WLL。 |

**方法**

|                                                              | 名称       | 说明               |
| ------------------------------------------------------------ | ---------- | ------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的加载项。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Autoload**    | 如果该属性为 **True**，则在 WPS 启动时，自动加载指定的加载项。位于 WPS 程序文件夹的 Startup 子文件夹中的加载项可以自动加载。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Compiled**    | 如果为 **True**，则指定的加载项是一个 WPS 加载项库 (WLL).如果为 **False**，则加载项为一个模板。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 类型的值，该值代表项在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Installed**   | 如果为 **True**，则指定的加载项已安装（即加载）。可在**“工具”**菜单的**“模板和加载项”**对话框中选择需要加载的加载项。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回加载项的名称。**String** 类型，只读。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **AddIn** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**        | 返回已安装加载项的位置。**String** 类型，只读。              |

**成员方法**

#### **AddIn.Delete**

删除指定的加载项。

**语法**

**express.Delete()**

*express*   一个代表 **AddIn** 对象的变量。

**成员属性**

#### **AddIn.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Autoload**

如果该属性为 **True**，则在 WPS 启动时，自动加载指定的加载项。位于 WPS 程序文件夹的 Startup 子文件夹中的加载项可以自动加载。**Boolean** 类型，只读。

**语法**

**express.Autoload**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例显示 WPS 启动时自动加载的每一加载项的名称。*/ function test() {    let blnFound = false    for (let i = 1; i <= Application.AddIns.Count; i++) {        if (Application.AddIns.Item(i).Autoload == true) {            alert(Application.AddIns.Item(i).Name)            blnFound = true        }    }    if (blnFound !== true) {        alert("No add-ins were loaded automatically.")    }}``` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例判断名为“Gallery.dot”的加载项是否自动加载。*/ function test() {     let addinLoop      for (let i = 1; i <= Application.AddIns.Count; i++) {         addinLoop = Application.AddIns.Item(i)         if (addinLoop.Name.toLowerCase().indexOf("gallery.dot") != -1) {             if (addinLoop.Autoload == true) {                 alert("Autoload")             }         }     } }``` |

#### **AddIn.Compiled**

如果为 **True**，则指定的加载项是一个 WPS 加载项库 (WLL).如果为 **False**，则加载项为一个模板。**Boolean** 类型，只读。

**语法**

**express.Compiled**

*express*   一个代表 **AddIn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例判断当前加载的 WLL 的数目。*/ function test() {     let count = 0     for (let i = 1; i <= Application.AddIns.Count; i++) {         if (Application.AddIns.Item(i).Compiled == true && Application.AddIns.Item(i).Installed == true) {             count++         }     }     alert(count + " WLL's are loaded") }``` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*如果第一个加载项是模板，本示例卸载该模板并打开它。*/ function test() {     if (Application.AddIns.Item(1).Compiled == false) {         Application.AddIns.Item(1).Installed = false         Documents.Open(Application.AddIns.Item(1).Path + Application.PathSeparator + Application.AddIns.Item(1).Name)     } }``` |

#### **AddIn.Creator**

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **AddIn.Index**

返回一个 **Long** 类型的值，该值代表项在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Installed**

如果为 **True**，则指定的加载项已安装（即加载）。可在**“工具”**菜单的**“模板和加载项”**对话框中选择需要加载的加载项。可读/写 **Boolean** 类型。

**语法**

**express.Installed**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

已卸载的加载项包含在 **AddIns** 集合中。若要从 AddIns 集合中删除模板或 WLL，可对 **AddIn** 对象应用 **Delete** 方法（加载项名称从**“模板和加载项”**对话框中删除）。若要卸载所有的模板和 WLL，可对 **AddIns** 集合应用 **Unload** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例卸载名为“Gallery.dot”的全局模板。*/ Application.AddIns.Item("Gallery.dot").Installed = false``` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例加载 FindAll.wll。*/ Application.AddIns.Item("C:\\Templates\\FindAll.wll").Installed = true ``` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*此示例加载 Custom.dot。*/ Application.AddIns.Item("C:\\Program Files\\Microsoft Office\\Templates\\Custom.dot").Installed = true``` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*如果 Dot1.dot 作为全局模板加载，本示例则在状态栏上显示一条消息。*/ function test() {     if (Application.AddIns.Item("Dot1.dot").Installed == true) {         StatusBar = "Dot1.dot is loaded"     } }``` |

#### **AddIn.Name**

返回加载项的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Parent**

返回一个 **Object** 类型值，该值代表指定 **AddIn** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **AddIn** 对象的变量。

#### **AddIn.Path**

返回已安装加载项的位置。**String** 类型，只读。

**语法**

**express.Path**

*express*   一个代表 **AddIn** 对象的变量。

**说明**

此路径不包括尾随字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

| 注释                                                         |
| ------------------------------------------------------------ |
| 可以使用 **PathSeparator** 属性建立 Web 地址，即使此地址中包含正斜杠 (/)，而 **PathSeparator** 属性默认使用反斜杠 (\)。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| ```/*本示例显示 AddIns 集合中第一个加载项的路径。*/ function test() {     if (Application.AddIns.Count >= 1) {         alert(Application.AddIns.Item(1).Path)     } }``` |

适用环境：web

适用平台：windows/linux