**Breaks** 



页面中分页符、分栏符和分节符的集合。使用 **Breaks** 集合及相关对象和属性可通过编程方式定义文档的页面版式。

**方法**

|                                                              | 名称     | 说明                              |
| ------------------------------------------------------------ | -------- | --------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Break** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象，该对象代表 WPS 应用程序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回 **Breaks** 集合中的项目数。**Long** 类型，只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表指定 **Breaks** 集合中的父对象。 |

**成员方法**

#### **Breaks.Item**

返回集合中的单个 **Break** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Breaks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 **Long**类型值。 |

**成员属性**

#### **Breaks.Application**

返回一个 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Breaks** 对象的变量。

#### **Breaks.Count**

返回 **Breaks** 集合中的项目数。**Long** 类型，只读。

**语法**

**express.Count**

*express*   一个代表 **Breaks** 对象的变量。

#### **Breaks.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Breaks** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

#### **Breaks.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Breaks** 集合中的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Breaks** 对象的变量。

适用环境：web

适用平台：windows/linux