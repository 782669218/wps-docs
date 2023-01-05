**TableStyles**



代表可应用于表格的样式。

**说明**

表格样式提供了一种为整个表格或数据透视图设置格式的方式。表格样式取代了用于为整个表格设置格式的现有的自动套用格式功能。

表格样式与自动套用格式在以下几个方面不同：

- 可以创建和重用自定义表格样式。
- 表格样式可处理主题。
- 如果更改文档主题的配色方案和/或字体方案，则将更改内置表格样式的外观。
- 当对象发生变化时，表格样式可以将样式重新应用于像数据透视表和表格之类的对象。该表格将记住应用于对象的样式，在添加、删除、隐藏和显示单元格时，表格将相应地进行重新显示。
- 表格样式在功能区中具有可见的用户界面。

**方法**

|                                                              | 名称     | 说明                                               |
| ------------------------------------------------------------ | -------- | -------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 创建新的 **TableStyle** 对象，并将其添加到集合中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。                             |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回集合中对象的数目。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读                                   |

**成员方法**

#### **TableStyles.Add**

创建新的 **TableStyle** 对象，并将其添加到集合中。

**语法**

**express.Add(TableStyleName)**

*express*   一个代表 **TableStyles** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**       |
| ---------------- | ------------- | ------------ | -------------- |
| *TableStyleName* | 必选          | **String**   | 表样式的名称。 |

**返回值**

TableStyle

#### **TableStyles.Item**

从集合中返回一个对象。

**语法**

**express.Item(index)**

*express*   一个代表 **TableStyles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

TableStyle

**成员属性**

#### **TableStyles.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **TableStyles** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **TableStyles.Count**

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

*express*   一个代表 **TableStyles** 对象的变量。

#### **TableStyles.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型

**语法**

**express.Creator**

*express*   一个代表 **TableStyles** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **TableStyles.Parent**

返回指定对象的父对象。只读

**语法**

**express.Parent**

*express*   一个代表 **TableStyles** 对象的变量。

适用环境：web

适用平台：windows/linux