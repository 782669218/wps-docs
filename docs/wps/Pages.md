#### **Pages**



文档中页面的集合。使用 **Pages** 集合及相关对象和属性可通过编程方式定义文档的页面版式。

**方法**

|                                                              | 名称     | 说明                                       |
| ------------------------------------------------------------ | -------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Page** 对象。返回Page值 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中的页数。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Pages** 对象的父对象。 |

**成员方法**

#### **Pages.Item**

返回集合中的单个 **Page** 对象。返回Page值

**语法**

**express.Item(Index)**

*express*   一个代表 **Pages** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**成员属性**

#### **Pages.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Pages** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

**示例**

| 示例代码复制                                     |
| ------------------------------------------------ |
| `返回一个代表 WPS 应用程序的 Application 对象。` |

#### **Pages.Count**

返回一个 **Long** 类型的值，该值代表集合中的页数。只读。

**语法**

**express.Count**

*express*   一个代表 **Pages** 对象的变量。

#### **Pages.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Pages** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Pages.Parent**

返回一个 **Object** 类型值，该值代表指定 **Pages** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Pages** 对象的变量。

适用环境：web

适用平台：windows/linux