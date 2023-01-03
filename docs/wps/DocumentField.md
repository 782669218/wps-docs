**DocumentField** 



代表文档中的一个公文域。

**方法**

|                                                              | 名称       | 说明           |
| ------------------------------------------------------------ | ---------- | -------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择指定的域。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**      | 如果为 **True**，则隐藏文档中的公文域。**Boolean**类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置指定对象的名称。可读/写   **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **DocumentField** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintOut**    | 该公文域对象是否可以打印 。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回指定对象的区域 Range。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**    | 如果为True，代表该公文域值只读。**Boolean**类型， 可读/写    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StoryType**   | 返回一个Long类型值，该值代表此公文域的类型。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 返回一个字符串，该字符串为此公文域中的内容。**String** 类型，可读/写。 |

**成员方法**

#### **DocumentField.Delete**

删除指定的域。

**语法**

**express.Delete(DeleteWithContent)**

*express*   一个代表 **DocumentField** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                          |
| ------------------- | ------------- | ------------ | --------------------------------- |
| *DeleteWithContent* | 可选          | **Boolean**  | 是否删除该域中的内容，默认为false |

#### **DocumentField.Select**

选择指定的域。

**语法**

**express.Select()**

*express*   一个代表 **DocumentField** 对象的变量。

**成员属性**

#### **DocumentField.Application**

返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/Application)对象。

**语法**

**express.Application**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **DocumentField** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **DocumentField.Hidden**

如果为 **True**，则隐藏文档中的公文域。**Boolean**类型，可读写。

**语法**

**express.Hidden**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.Name**

返回或设置指定对象的名称。可读/写   **String** 类型。

**语法**

**express.Name**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.Parent**

返回一个 **Object** 类型值，该值代表指定 **DocumentField** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.PrintOut**

该公文域对象是否可以打印 。

**语法**

**express.PrintOut**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.Range**

返回指定对象的区域 Range。

**语法**

**express.Range**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.ReadOnly**

如果为True，代表该公文域值只读。**Boolean**类型， 可读/写

**语法**

**express.ReadOnly**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.StoryType**

返回一个Long类型值，该值代表此公文域的类型。只读。

**语法**

**express.StoryType**

*express*   一个代表 **DocumentField** 对象的变量。

#### **DocumentField.Value**

返回一个字符串，该字符串为此公文域中的内容。**String** 类型，可读/写。

**语法**

**express.Value**

*express*   一个代表 **DocumentField** 对象的变量。

适用环境：web

适用平台：windows/linux