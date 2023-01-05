**DocumentFields** 



代表一个公文域集合。**DocumentField** 对象是 **DocumentFields** 集合的成员。**DocumentFields** 集合中包含当前在 WPS 中打开的所有 **DocumentField** 对象。

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Exists** | 判断是否存在该名称的公文域，如果存在范围True。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 返回一个 **DocumentField**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/DocumentField/DocumentField%20.htm#jsObject_DocumentField)对象，该对象通过索引来确定。 |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**          | 返回一个**Long**类型的值，代表该对象中**DocumentField**对象的数量，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultSorting** | 设置DocumentFields集合的排序方式，可以通过枚举值 [**WpsBookmarkSortBy** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WpsBookmarkSortBy%20%E6%9E%9A%E4%B8%BE.html)指定。 可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertionMode**  | 通过Boolean类型变量设置公文域对象的插入模式， 可读/写。      |

**成员方法**

#### **DocumentFields.Add**

**语法**

**express.Add(Name, Range, Hidden, PrintOut, ReadOnly)**

*express*   一个代表 **DocumentFields** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                  |
| ---------- | ------------- | ------------ | ----------------------------------------- |
| *Name*     | 必选          | **String**   | 要添加的公文域名称                        |
| *Range*    | 可选          | **Variant**  | 代表一个Range对象，表示该公文域添加的位置 |
| *Hidden*   | 可选          | **Variant**  | 该公文域添加后是否显示                    |
| *PrintOut* | 可选          | **Variant**  | 在文档打印时是否打印该公文域              |
| *ReadOnly* | 可选          | **Variant**  | 该公文域是否为只读类型                    |

#### **DocumentFields.Exists**

判断是否存在该名称的公文域，如果存在范围True。

**语法**

**express.Exists(Name)**

*express*   一个代表 **DocumentFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Name*   | 必选          | **String**   | 公文域得名称。 |

#### **DocumentFields.Item**

返回一个 **DocumentField**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/DocumentField/DocumentField%20.htm#jsObject_DocumentField)对象，该对象通过索引来确定。

**语法**

**express.Item(Index)**

*express*   一个代表 **DocumentFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 要获取的公文域索引值 |

**成员属性**

#### **DocumentFields.Count**

返回一个**Long**类型的值，代表该对象中**DocumentField**对象的数量，只读。

**语法**

**express.Count**

*express*   一个代表 **DocumentFields** 对象的变量。

#### **DocumentFields.DefaultSorting**

设置DocumentFields集合的排序方式，可以通过枚举值 [**WpsBookmarkSortBy** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WpsBookmarkSortBy%20%E6%9E%9A%E4%B8%BE.html)指定。 可读/写

**语法**

**express.DefaultSorting**

*express*   一个代表 **DocumentFields** 对象的变量。

#### **DocumentFields.InsertionMode**

通过Boolean类型变量设置公文域对象的插入模式， 可读/写。

**语法**

**express.InsertionMode**

*express*   一个代表 **DocumentFields** 对象的变量。

适用环境：web

适用平台：windows/linux