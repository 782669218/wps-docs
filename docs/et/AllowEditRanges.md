**AllowEditRanges**



所有 [AllowEditRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/AllowEditRange/AllowEditRange%20.htm#jsObject_AllowEditRange)对象的集合，这些对象代表受保护工作表上的可编辑单元格。

**方法**

|                                                              | 名称    | 说明                                                         |
| ------------------------------------------------------------ | ------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 在受保护的工作表中添加可编辑的单元格区域。返回 [AllowEditRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/AllowEditRange/AllowEditRange%20.htm#jsObject_AllowEditRange)对象。 |

**属性**

|                                                              | 名称      | 说明                                           |
| ------------------------------------------------------------ | --------- | ---------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count** | 返回一个 **Long** 值，它代表集合中对象的数量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**  | 从集合中返回一个对象。                         |

**成员方法**

#### **AllowEditRanges.Add**

在受保护的工作表中添加可编辑的单元格区域。返回 [AllowEditRange ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/AllowEditRange/AllowEditRange%20.htm#jsObject_AllowEditRange)对象。

**语法**

**express.Add(Title, Range, Password)**

*express*   一个代表 **AllowEditRanges** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                           |
| ---------- | ------------- | ------------ | ---------------------------------- |
| *Title*    | 必选          | **String**   | 单元格区域的标题。                 |
| *Range*    | 必选          | **Object**   | Range 对象。允许编辑的单元格区域。 |
| *Password* | 可选          | **Variant**  | 单元格区域的密码。                 |

**成员属性**

#### **AllowEditRanges.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **AllowEditRanges** 对象的变量。

#### **AllowEditRanges.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **AllowEditRanges** 对象的变量。

适用环境：web

适用平台：windows/linux