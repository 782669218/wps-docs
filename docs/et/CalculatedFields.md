**CalculatedFields**



由 **PivotField** 对象组成的集合，这些对象代表指定数据透视表中的所有计算字段。

**方法**

|                                                              | 名称     | 说明                                         |
| ------------------------------------------------------------ | -------- | -------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 创建新的计算字段。返回 **PivotField** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **CalculatedFields.Add**

创建新的计算字段。返回 **PivotField** 对象。

**语法**

**express.Add(Name, Formula, UseStandardFormula)**

*express*   一个代表 **CalculatedFields** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*               | 必选          | **String**   | 字段的名称。                                                 |
| *Formula*            | 必选          | **String**   | 字段的公式。                                                 |
| *UseStandardFormula* | 可选          | **Variant**  | 为实现向上兼容，该值默认为 False。对于作为字段名称的任何参数中所含的字符串来说，该值为 True，并将被解释为已按标准美国英语进行了格式设置，而非本地设置。 |

**说明**

返回值
一个代表新计算字段的 **PivotField**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例向第一张工作表上的第一张数据透视表添加计算字段。*/ Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Add("PxS", "= Product * Sales")` |

#### **CalculatedFields.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **CalculatedFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**说明**

返回值包含在集合中的一个 **PivotField** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为第一个计算字段设置公式。*/ Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Item(1).Formula = "=Revenue - Cost"` |

**成员属性**

#### **CalculatedFields.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 Application 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CalculatedFields** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   if(Application.ActiveWorkbook.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **CalculatedFields.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **CalculatedFields** 对象的变量。

#### **CalculatedFields.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CalculatedFields** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **CalculatedFields.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CalculatedFields** 对象的变量。

适用环境：web

适用平台：windows/linux