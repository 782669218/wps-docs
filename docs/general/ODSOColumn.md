#### **ODSOColumn**



代表数据源中的一个字段。**ODSOColumn** 对象是 **ODSOColumns** 集合的成员。

**说明**

**ODSOColumns** 集合包括邮件合并数据源中的所有数据字段（例如，姓名、地址和城市）。

不能向 **ODSOColumns** 集合添加字段。**ODSOColumns** 集合自动包括数据源中的所有数据字段。

使用 Columns (*index*) 返回单个 **ODSOColumn** 对象，其中 *index* 是数据字段名或索引号。索引号表示数据字段在邮件合并数据源中的位置。

 

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **ODSOColum**n 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **ODSOColumn** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 获取一个 **Long** 类型的值，代表集合中 **ODSOColumn** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 获取邮件合并数据源中的数据字段的名称。只读。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **ODSOColumn** 对象的 **Parent** 对象。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 获取邮件合并数据源中的数据字段的值。只读。                   |

**成员属性**

#### **ODSOColumn.Application**

获取一个 **Application** 对象，代表 **ODSOColum**n 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **ODSOColumn** 对象的变量。

#### **ODSOColumn.Creator**

获取一个 32 位整数，指示创建 **ODSOColumn** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **ODSOColumn** 对象的变量。

#### **ODSOColumn.Index**

获取一个 **Long** 类型的值，代表集合中 **ODSOColumn** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **ODSOColumn** 对象的变量。

#### **ODSOColumn.Name**

获取邮件合并数据源中的数据字段的名称。只读。

**语法**

**express.Name**

*express*   一个代表 **ODSOColumn** 对象的变量。

#### **ODSOColumn.Parent**

获取 **ODSOColumn** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ODSOColumn** 对象的变量。

#### **ODSOColumn.Value**

获取邮件合并数据源中的数据字段的值。只读。

**语法**

**express.Value**

*express*   一个代表 **ODSOColumn** 对象的变量。

适用环境：web

适用平台：windows/linux