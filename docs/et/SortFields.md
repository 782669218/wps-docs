**SortFields**



**SortFields**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/SortFields/SortFields%20.htm#jsObject_SortFields)集合是 **SortField**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/SortField/SortField%20.htm#jsObject_SortField)对象的集合。开发人员可以使用该集合存储工作簿、列表和自动筛选的排序状态。

**说明**

该对象具有添加、计数、排序和删除 **SortField** 对象的属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     Application.ActiveSheet.Sort.SortFields.Add(Range("A1"),null,xlDescending)      Application.ActiveSheet.Sort.SortFields.Add(Range("B1"),null,xlDescending)   }` |

**方法**

|                                                              | 名称      | 说明                                                         |
| ------------------------------------------------------------ | --------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**   | 创建新的排序字段，并返回一个 **SortFields** 对象。返回SortField值 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear** | 清除所有 **SortFields** 对象。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**  | 返回一个 **SortField** 对象，该对象代表可以存储在工作簿中的项的集合。只读。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回集合中对象的数目。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **SortFields.Add**

创建新的排序字段，并返回一个 **SortFields** 对象。返回SortField值

**语法**

**express.Add(Key, SortOn, Order, CustomOrder, DataOption)**

*express*   一个代表 **SortFields** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                       |
| ------------- | ------------- | ------------ | ------------------------------ |
| *Key*         | 必选          | **Range**    | 指定用于排序的键值。           |
| *SortOn*      | 可选          | **Variant**  | 要进行排序的字段。             |
| *Order*       | 可选          | **Variant**  | 指定排序次序。                 |
| *CustomOrder* | 可选          | **Variant**  | 指定是否应使用自定义排序次序。 |
| *DataOption*  | 可选          | **Variant**  | 指定数据选项。                 |

#### **SortFields.Clear**

清除所有 **SortFields** 对象。

**语法**

**express.Clear()**

*express*   一个代表 **SortFields** 对象的变量。

#### **SortFields.Item**

返回一个 **SortField** 对象，该对象代表可以存储在工作簿中的项的集合。只读。

**语法**

**express.Item(Index)**

*express*   一个代表 **SortFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Index*  | 必选          | **Variant**  | 排序字段的索引值。 |

**成员属性**

#### **SortFields.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **SortFields** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **SortFields.Count**

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

*express*   一个代表 **SortFields** 对象的变量。

#### **SortFields.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **SortFields** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **SortFields.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SortFields** 对象的变量。

适用环境：web

适用平台：windows/linux