**PivotCaches**



代表工作簿中数据透视表的内存缓存的集合。

**说明**

每一个内存缓存都由一个 **PivotCache** 对象代表。

使用 **PivotCaches** 方法可返回 **PivotCaches** 集合。下例为活动工作簿中的所有内存缓存设置 **RefreshOnFileOpen** 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){for (let i=1;i<=ActiveWorkbook.PivotCaches.Count;i++) {     ActiveWorkbook.PivotCaches().Item(i).RefreshOnFileOpen = true } }` |

**方法**

|                                                              | 名称       | 说明                   |
| ------------------------------------------------------------ | ---------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Create** | 创建新的 PivotCache。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **PivotCaches.Create**

创建新的 PivotCache。

**语法**

**express.Create(SourceType, SourceData, Version)**

*express*   一个代表 **PivotCaches** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**               | **说明**                                                     |
| ------------ | ------------- | -------------------------- | ------------------------------------------------------------ |
| *SourceType* | 必选          | **XlPivotTableSourceType** | SourceType 可以为以下 xlPivotTableSourceType 常量之一：xlConsolidation、xlDatabase 或 xlExternal。 |
| *SourceData* | 可选          | **Variant**                | 新数据透视表缓存的数据。                                     |
| *Version*    | 可选          | **Variant**                | 数据透视表的版本。可以为 xlPivotTableVersionList 常量之一。  |

**返回值**

PivotCache

**说明**

使用此方法创建 PivotCache 时，不支持以下两个 **xlPivotTableSourceType** 常量：**xlPivotTable** 和 **xlScenario**。如果提供这两个常量之一，将返回运行时错误。

如果 *SourceType* 不为 **xlExternal**，则 *SourceData* 参数为必需。该参数可以为 **Range** 对象（当 *SourceType* 为 **xlConsolidation** 或 **xlDatabase** 时），或为 ET 工作簿连接对象（当 *SourceType* 为 **xlExternal** 时）。

如果不提供数据透视表的版本，则版本默认为 **xlPivotTableVersion12**。不允许使用 **xlPivotTableVersionCurrent** 常量，如果提供该常量，将返回运行时错误。

#### **PivotCaches.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **PivotCaches** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 PivotCache 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例刷新第一个缓存。*/  ActiveWorkbook.PivotCaches().Item(1).Refresh()` |

**成员属性**

#### **PivotCaches.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotCaches** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotCaches.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **PivotCaches** 对象的变量。

#### **PivotCaches.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotCaches** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotCaches.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotCaches** 对象的变量。

适用环境：web

适用平台：windows/linux