**TableStyleElements**



代表表格样式元素。

**说明**

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行、最后一列或总计行是表格的元素，表格样式可以规定标题行的填充色为蓝色，最后一列为红色。

表格中表格样式元素的格式设置可在适用于该元素的表格样式中指定。**XlTableStyleElementType** 枚举包含可供使用的表格样式元素的类型。

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回集合中对象的数目。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **TableStyleElements.Item**

从集合中返回一个对象。

**语法**

**express.Item(index)**

*express*   一个代表 **TableStyleElements** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**                | **说明**     |
| -------- | ------------- | --------------------------- | ------------ |
| *index*  | 必选          | **XlTableStyleElementType** | 表样式元素。 |

**返回值**

TableStyleElement

**成员属性**

#### **TableStyleElements.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **TableStyleElements** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **TableStyleElements.Count**

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

*express*   一个代表 **TableStyleElements** 对象的变量。

#### **TableStyleElements.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **TableStyleElements** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **TableStyleElements.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **TableStyleElements** 对象的变量。

适用环境：web

适用平台：windows/linux