#### **FileConverters**



**FileConverter** 对象的集合，这些对象代表可用于打开和保存文件的所有文件转换器。

**方法**

|                                                              | 名称     | 说明                                      |
| ------------------------------------------------------------ | -------- | ----------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **FileConverter** 对象。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConvertMacWordChevrons** | 控制是否将 V 型字符 (? ?) 中的文字转化成合并域。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                  | 返回一个 **Long** 类型的值，该值代表集合中文件转换器的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | 返回一个 **Object** 类型值，该值代表指定 **FileConverters** 对象的父对象。 |

**成员方法**

#### **FileConverters.Item**

返回集合中的单个 **FileConverter** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **FileConverters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**返回值**

FileConverter

**成员属性**

#### **FileConverters.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **FileConverters** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **FileConverters.ConvertMacWordChevrons**

控制是否将 V 型字符 (? ?) 中的文字转化成合并域。**Long** 类型，可读写。

**语法**

**express.ConvertMacWordChevrons**

*express*   一个代表 **FileConverters** 对象的变量。

**说明**

**ConvertMacWordChevrons** 属性可以是任何 **WdChevronConvertRule** 常量。

适用于 Macintosh 的 WPS 4.0 和 5.x 版本的文档用 V 型字符指出邮件合并域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 ConvertMacWordChevrons 属性，将 V 型字符中的文本转换成邮件合并域，然后打开名为“Mac WPS Document”的文档。*/ function test() { FileConverters.ConvertMacWordChevrons = wdAlwaysConvert Documents.Open("C:\\Documents\\Mac  WPS Document") }` |

#### **FileConverters.Count**

返回一个 **Long** 类型的值，该值代表集合中文件转换器的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **FileConverters** 对象的变量。

#### **FileConverters.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FileConverters** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **FileConverters.Parent**

返回一个 **Object** 类型值，该值代表指定 **FileConverters** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **FileConverters** 对象的变量。

适用环境：web

适用平台：windows/linux