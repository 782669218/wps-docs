#### **ODSOFilter**



表示应用于附加的邮件合并数据源的筛选。**ODSOFilter** 对象为 **ODSOFilters** 对象的成员。

**说明**

每个筛选器均为查询字符串中的一行。使用 **Column**、**Comparison**、**CompareTo** 和 **Conjunction** 属性可返回或设置数据源查询条件。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **ODSOFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Column**      | 获取或设置一个 **String** 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CompareTo**   | 获取或设置一个 **String** 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comparison**  | 获取或设置一个代表 **Column** 和 **CompareTo** 的比较方式的 **MsoFilterComparison** 常量。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Conjunction** | 获取或设置一个 **MsoFilterConjunction** 常量，该常量代表 **ODSOFilters** 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **ODSOFilter** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 获取一个 **Long** 类型的值，代表集合中 **ODSOFilter** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **ODSOFilter** 对象的 **Parent** 对象。只读。           |

**成员属性**

#### **ODSOFilter.Application**

获取一个 **Application** 对象，代表 **ODSOFilter** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **ODSOFilter** 对象的变量。

#### **ODSOFilter.Column**

获取或设置一个 **String** 类型的值，该值代表邮件合并数据源中要用于筛选的字段名。可读/写。

**语法**

**express.Column**

*express*   一个代表 **ODSOFilter** 对象的变量。

**示例**

#### **ODSOFilter.CompareTo**

获取或设置一个 **String** 类型的值，该值代表查询筛选条件中要比较的文本。可读/写。

**语法**

**express.CompareTo**

*express*   一个代表 **ODSOFilter** 对象的变量。

**示例**

#### **ODSOFilter.Comparison**

获取或设置一个代表 **Column** 和 **CompareTo** 的比较方式的 **MsoFilterComparison** 常量。可读/写。

**语法**

**express.Comparison**

*express*   一个代表 **ODSOFilter** 对象的变量。

**示例**

#### **ODSOFilter.Conjunction**

获取或设置一个 **MsoFilterConjunction** 常量，该常量代表 **ODSOFilters** 对象中的某个筛选条件与其中的其他筛选条件的关系。可读/写。

**语法**

**express.Conjunction**

*express*   一个代表 **ODSOFilter** 对象的变量。

**示例**

#### **ODSOFilter.Creator**

获取一个 32 位整数，指示创建 **ODSOFilter** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **ODSOFilter** 对象的变量。

#### **ODSOFilter.Index**

获取一个 **Long** 类型的值，代表集合中 **ODSOFilter** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **ODSOFilter** 对象的变量。

#### **ODSOFilter.Parent**

获取 **ODSOFilter** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ODSOFilter** 对象的变量。

适用环境：web

适用平台：windows/linux