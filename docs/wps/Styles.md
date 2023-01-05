#### **Styles**



**Style** 对象的集合，这些对象代表文档中的内置样式和用户定义的样式。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **HeadingStyle** 对象，该对象代表添加到文档中的新标题样式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Style** 对象。                            |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中样式的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Styles** 对象的父对象。 |

**成员方法**

#### **Styles.Add**

返回一个 **HeadingStyle** 对象，该对象代表添加到文档中的新标题样式。

**语法**

**express.Add(Style, Level)**

*express*   一个代表 **Styles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Style*  | 必选          | **Variant**  | 要添加的样式。要指定该参数，可以使用该样式的字符串名称或 **Style** 对象。 |
| *Level*  | 必选          | **Integer**  | 代表标题级别的数值。                                         |

**说明**

在编译目录或图表目录时，都会包含该新标题样式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开始处添加一个目录，然后向用于创建目录的样式列表添加 Title 样式*/ let myToc = Application.ActiveDocument.TablesOfContents.Add(ActiveDocument.Range(0, 0), true, 1, 3) myToc.HeadingStyles.Add("title", 2)` |

#### **Styles.Item**

返回集合中的单个 **Style** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Styles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 **Long** 类型值，或代表单个对象名称的 **String** 类型值。 |

**成员属性**

#### **Styles.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Styles** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Styles.Count**

返回一个 **Long** 类型的值，该值代表集合中样式的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **Styles** 对象的变量。

**说明**

#### **Styles.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Styles** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Styles.Parent**

返回一个 **Object** 类型值，该值代表指定 **Styles** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Styles** 对象的变量。

适用环境：web

适用平台：windows/linux