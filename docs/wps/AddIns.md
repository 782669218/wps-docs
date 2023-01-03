**AddIns** 对象

卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 **AddIns** 集合中相应的加载项。

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 返回一个 **AddIn** 对象，该对象表示一个添加到可用加载项列表中的加载项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 返回集合中的单个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unload** | 卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 **AddIns** 集合中相应的加载项。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回 **AddIns** 集合中 **AddIn** 对象的数目。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表 **Addins** 集合中的父对象。通常为 **Application** 对象。 |

**成员方法**

#### **AddIns.Add**

返回一个 **AddIn** 对象，该对象表示一个添加到可用加载项列表中的加载项。

**语法**

**express.Add(FileName, Install)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName* | 必选          | **String**   | 模板或 WLL 的路径。                                          |
| *Install*  | 可选          | **Variant**  | 如果为 True，则安装加载项；如果为 False，则将加载项添加到加载项列表中，但不进行安装。默认值为 True。 |

**说明**

可以使用加载项的 **Installed** 属性查看该项是否已经安装。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例安装名为 MyFax.dot 的模板并将其添加至“模板和加载项”对话框中的加载项列表中。 */ function test() {  /* For this example to work correctly, verify that the        path is correct and the file exists. */      Application.AddIns.Add("C:\\Program Files\\Microsoft Office\\Templates\\Letters Faxes\\MyFax.dot",true)  }` |

#### **AddIns.Item**

返回集合中的单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个加载项。Index 可以是代表加载项在集合中的排序位置的 Long 类型值，或者是代表单个加载项名称的 String 类型值。 |

#### **AddIns.Unload**

卸载所有已加载的加载项，并根据 *RemoveFromList* 参数值删除 **AddIns** 集合中相应的加载项。

**语法**

**express.Unload(RemoveFromList)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**       | **说明**                                                     |
| ---------------- | ------------- | ------------------ | ------------------------------------------------------------ |
| *RemoveFromList* | 必选          | **RemoveFromList** | 如果为 True，则从 AddIns 集合中删除已卸载的加载项（从“模板和加载项”对话框中删除名称）。如果为 False，则保留集合中已卸载的加载项。如果已卸载的加载项的 Autoload 属性返回 True，则 Unload 无法从 AddIns 集合中删除该加载项，而不考虑 RemoveFromList 的值是什么。 |

**说明**

要卸载单个模板或 WLL，请将 **AddIn** 对象的 **Installed** 属性设置为 **False**。要从 **AddIns** 集合中删除单个模板或 WLL，请将 **Delete** 方法应用于 **AddIn** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例卸载“模板和加载项”对话框中列出的所有加载项。而加载项的名称还保留在 AddIns 集合中。*/ function test() {     if (Application.AddIns.Count > 0) {         Application.AddIns.Unload(false)     } }` |

**成员属性**

#### **AddIns.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **AddIns** 对象的变量。

#### **AddIns.Count**

返回 **AddIns** 集合中 **AddIn** 对象的数目。**Long** 类型，只读。

**语法**

**express.Count**

*express*   一个代表 **AddIns** 对象的变量。

#### **AddIns.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **AddIns** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **AddIns.Parent**

返回一个 **Object** 类型的值，该值代表 **Addins** 集合中的父对象。通常为 **Application** 对象。

**语法**

**express.Parent**

*express*   一个代表 **AddIns** 对象的变量。

适用环境：web

适用平台：windows/linux