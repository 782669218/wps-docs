#### **XMLNamespaces**



**XMLNamespace** 对象的集合，该集合代表架构库中架构的整个集合。

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**             | 返回一个 **XMLNamespace** 对象，该对象代表添加到架构库中并在 WPS 中对用户可用的架构。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InstallManifest** | 将指定的 XML 扩展包安装在用户计算机上，令一个或更多用户能够使用 XML 智能文档解决方案。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**            | 返回集合中的单个 **XMLNamespace** 对象。                     |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中 XML 命名空间的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **XMLNamespaces** 对象的父对象。 |

**成员方法**

#### **XMLNamespaces.Add**

返回一个 **XMLNamespace** 对象，该对象代表添加到架构库中并在 WPS 中对用户可用的架构。

**语法**

**express.Add(Path, NamespaceURI, Alias, InstallForAllUsers)**

*express*   一个代表 **XMLNamespaces** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Path*               | 必选          | **String**   | 架构的路径和文件名。可以为本地文件路径、网络路径或 Internet 地址。 |
| *NamespaceURI*       | 可选          | **String**   | 架构中指定的命名空间“统一资源标识符”。NamespaceURI 参数区分大小写，因此必须严格按照架构中所指定的那样进行拼写。 |
| *Alias*              | 可选          | **String**   | 在“模板和加载项”对话框中的“架构”选项卡上显示的架构名称。     |
| *InstallForAllUsers* | 可选          | **Boolean**  | 如果登录到计算机上的所有用户都能访问和使用新架构，则为 True。默认值为 False。 |

**返回值**

XMLNamespace

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将指定的架构添加到架构库，然后将其附加到活动文档中。本示例假设在指定的路径中已有一个名为 simplesample.xsd 的架构。*/ function AddSchema() {     let objSchema = Application.XMLNamespaces.Add("C:\\schemas\\simplesample.xsd")            objSchema.AttachToDocument(ActiveDocument) }  ` |

#### **XMLNamespaces.InstallManifest**

将指定的 XML 扩展包安装在用户计算机上，令一个或更多用户能够使用 XML 智能文档解决方案。

**语法**

**express.InstallManifest(Path, InstallForAllUsers)**

*express*   一个代表 **XMLNamespaces** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Path*               | 必选          | **String**   | XML 扩展包的路径和文件名。                                   |
| *InstallForAllUsers* | 可选          | **Boolean**  | 如果为 True，则安装 XML 扩展包并允许该计算机上的所有用户使用。如果为 False，则 XML 扩展包仅限于当前用户使用。默认值为 False。 |

**说明**

下面的代码示例将 SimpleSample 智能文档解决方案安装到用户计算机上并使其仅用于当前用户。

| 注释                                                         |
| ------------------------------------------------------------ |
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Smart Document SDK。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.XMLNamespaces.InstallManifest("http://smartdocuments/simplesample/manifest.xml")` |

#### **XMLNamespaces.Item**

返回集合中的单个 **XMLNamespace** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **XMLNamespaces** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

XMLNamespace

**成员属性**

#### **XMLNamespaces.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **XMLNamespaces** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLNamespaces.Count**

返回一个 **Long** 类型的值，该值代表集合中 XML 命名空间的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **XMLNamespaces** 对象的变量。

#### **XMLNamespaces.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLNamespaces** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLNamespaces.Parent**

返回一个 **Object** 类型值，该值代表指定 **XMLNamespaces** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLNamespaces** 对象的变量。

适用环境：web

适用平台：windows/linux