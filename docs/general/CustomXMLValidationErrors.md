#### **CustomXMLValidationErrors**



代表 **CustomXMLValidationError** 对象的集合。

**方法**

|                                                              | 名称    | 说明                                                         |
| ------------------------------------------------------------ | ------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 将包含 XML 验证错误的 **CustomXMLValidationError** 对象添加到 **CustomXMLValidationErrors** 集合。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个代表 **CustomXMLValidationErrors** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **CustomXMLValidationErrors** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **CustomXMLValidationErrors** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 获取 **CustomXMLValidationErrors** 集合中的一个 **CustomXMLValidationError** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **CustomXMLValidationErrors** 对象的 **Parent** 对象。只读。 |

**成员方法**

#### **CustomXMLValidationErrors.Add**

将包含 XML 验证错误的 **CustomXMLValidationError** 对象添加到 **CustomXMLValidationErrors** 集合。

**语法**

**express.Add(Node, ErrorName, ErrorText, ClearedOnUpdate)**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**      | **说明**                                                     |
| ----------------- | ------------- | ----------------- | ------------------------------------------------------------ |
| *Node*            | 必选          | **CustomXMLNode** | 代表其中发生了错误的节点。                                   |
| *ErrorName*       | 必选          | **String**        | 包含错误的名称。                                             |
| *ErrorText*       | 可选          | **String**        | 包含描述性错误文本。                                         |
| *ClearedOnUpdate* | 可选          | **Boolean**       | 指定在更正并更新了 XML 后是否要从 CustomXMLValidationErrors 集合中清除错误。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将一个错误消息添加到集合。 function test() {     let coreNode = Application.ActivePresentation.CustomXMLParts.Item(1).SelectSingleNode("/ns1:coreProperties[1]")     Application.ActivePresentation.CustomXMLParts.Item(1).Errors.Add(coreNode , "ValidationError", "To add content to this stream, it must be valid, well-formed XML.", true) }` |

**成员属性**

#### **CustomXMLValidationErrors.Application**

获取一个代表 **CustomXMLValidationErrors** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

#### **CustomXMLValidationErrors.Count**

获取一个 **Long** 类型的值，指示 **CustomXMLValidationErrors** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

#### **CustomXMLValidationErrors.Creator**

获取一个 32 位整数，指示创建 **CustomXMLValidationErrors** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

#### **CustomXMLValidationErrors.Item**

获取 **CustomXMLValidationErrors** 集合中的一个 **CustomXMLValidationError** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

#### **CustomXMLValidationErrors.Parent**

获取 **CustomXMLValidationErrors** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLValidationErrors** 对象的变量。

适用环境：web

适用平台：windows/linux