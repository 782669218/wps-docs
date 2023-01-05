#### **ODSOFilters**



代表应用于附加到邮件合并发布内容的数据源的所有筛选器。**ODSOFilters** 对象由 **ODSOFilter** 对象组成。

**说明**

**方法**

|                                                              | 名称       | 说明                                                    |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 在 **ODSOFilters** 集合中新添一个筛选。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 从 **ODSOFilters** 集合中删除筛选对象。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 代表 **ODSOFilters** 集合中的一个 **ODSOFilter** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Parent** | 获取 **ODSOFilters** 对象的 **Parent** 对象。只读。     |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **ODSOFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **ODSOFilters** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **ODSOFilters** 对象时所使用的应用程序。只读。 |

**成员方法**

#### **ODSOFilters.Add**

在 **ODSOFilters** 集合中新添一个筛选。

**语法**

**express.Add(Column, Comparison, Conjunction, bstrCompareTo, DeferUpdate)**

*express*   一个代表 **ODSOFilters** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**             | **说明**                                                     |
| --------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *Column*        | 必选          | **String**               | 数据源中的表名称。                                           |
| *Comparison*    | 必选          | **MsoFilterComparison**  | 筛选表中数据的方式。                                         |
| *Conjunction*   | 必选          | **MsoFilterConjunction** | 确定此筛选与 ODSOFilters 对象中其他筛选的关联方式。          |
| *bstrCompareTo* | 可选          | **String**               | 如果 Comparison 参数不是 msoFilterComparisonIsBlank 或 msoFilterComparisonIsNotBlank，则表中的数据与该字符串进行比较。 |
| *DeferUpdate*   | 可选          | **Boolean**              | 指定是否延迟对筛选器的更新。默认值为 False。                 |

#### **ODSOFilters.Delete**

从 **ODSOFilters** 集合中删除筛选对象。

**语法**

**express.Delete(Index, DeferUpdate)**

*express*   一个代表 **ODSOFilters** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                     |
| ------------- | ------------- | ------------ | -------------------------------------------- |
| *Index*       | 必选          | **Long**     | 要删除的筛选的编号。                         |
| *DeferUpdate* | 可选          | **Boolean**  | 指定是否延迟对筛选器的更新。默认值为 False。 |

#### **ODSOFilters.Item**

代表 **ODSOFilters** 集合中的一个 **ODSOFilter** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ODSOFilters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**     |
| -------- | ------------- | ------------ | ------------ |
| *Index*  | 必选          | **Long**     | 项目的编号。 |

**示例**

#### **ODSOFilters.Parent**

获取 **ODSOFilters** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent()**

*express*   一个代表 **ODSOFilters** 对象的变量。

**成员属性**

#### **ODSOFilters.Application**

获取一个 **Application** 对象，代表 **ODSOFilters** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **ODSOFilters** 对象的变量。

#### **ODSOFilters.Count**

获取一个 **Long** 类型的值，指示 **ODSOFilters** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **ODSOFilters** 对象的变量。

#### **ODSOFilters.Creator**

获取一个 32 位整数，指示创建 **ODSOFilters** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **ODSOFilters** 对象的变量。

适用环境：web

适用平台：windows/linux