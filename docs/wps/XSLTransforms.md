#### **XSLTransforms**



**XSLTransform** 对象的集合，这些对象代表特定 XML 命名空间的所有可扩展样式表语言转换 (XSLT)。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **XSLTransform** 对象，该对象代表添加到指定架构的 XSLT 集合的“可扩展样式表语言转换”(Extensible Stylesheet Language Transformation, XSLT)。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的 **XSLTransform** 对象。                         |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中 XSLTransform 的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **XSLTransforms** 对象的父对象。 |

**成员方法**

#### **XSLTransforms.Add**

返回一个 **XSLTransform** 对象，该对象代表添加到指定架构的 XSLT 集合的“可扩展样式表语言转换”(Extensible Stylesheet Language Transformation, XSLT)。

**语法**

**express.Add(Location, Alias, InstallForAllUsers)**

*express*   一个代表 **XSLTransforms** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Location*           | 必选          | **String**   | XSLT 的路径和文件名。可以为本地文件路径、网络路径或 Internet 地址。 |
| *Alias*              | 可选          | **String**   | 架构库中显示的 XSLT 名称。                                   |
| *InstallForAllUsers* | 可选          | **Boolean**  | 如果登录到计算机上的所有用户都能访问和使用新架构，则为 True。默认值为 False。 |

**返回值**

XSLTransform

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下代码将一个架构添加到架构库中，然后将 XSLT 添加到该新添加的架构中。*/ function AddXSLT() {     let objSchema = Application.XMLNamespaces.Item("SimpleSample")     let objTransform = objSchema.XSLTransforms.Add("c:\\schemas\\simplesample.xsl") }` |

#### **XSLTransforms.Item**

返回集合中的 **XSLTransform** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **XSLTransforms** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

XSLTransform

**成员属性**

#### **XSLTransforms.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **XSLTransforms** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XSLTransforms.Count**

返回一个 **Long** 类型的值，该值代表集合中 XSLTransform 的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **XSLTransforms** 对象的变量。

#### **XSLTransforms.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XSLTransforms** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XSLTransforms.Parent**

返回一个 **Object** 类型值，该值代表指定 **XSLTransforms** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XSLTransforms** 对象的变量。

适用环境：web

适用平台：windows/linux