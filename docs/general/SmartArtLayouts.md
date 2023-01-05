#### **SmartArtLayouts**



代表 Smart Art 布局图表的集合。

**说明**

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码更改 WPP 中的 Smart Art 图表的图表样式。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 检索位于指定索引处或具有指定的唯一 Id 的 **SmartArtLayout** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **SmartArtLayouts** 对象的容器应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 检索包含在 **SmartArtLayouts** 集合中的 **SmartArtLayout** 对象数的计数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **SmartArtLayouts** 对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回调用对象。只读。                                         |

**成员方法**

#### **SmartArtLayouts.Item**

检索位于指定索引处或具有指定的唯一 Id 的 **SmartArtLayout** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **SmartArtLayouts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 指定一个代表索引的整数或一个代表 SmartArtLayout 对象的位置的字符串。 |

**返回值**

SmartArtLayout

**成员属性**

#### **SmartArtLayouts.Application**

获取一个 **Application** 对象，该对象代表 **SmartArtLayouts** 对象的容器应用程序。只读。

**语法**

**express.Application**

*express*   一个代表 **SmartArtLayouts** 对象的变量。

#### **SmartArtLayouts.Count**

检索包含在 **SmartArtLayouts** 集合中的 **SmartArtLayout** 对象数的计数。只读。

**语法**

**express.Count**

*express*   一个代表 **SmartArtLayouts** 对象的变量。

#### **SmartArtLayouts.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtLayouts** 对象的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SmartArtLayouts** 对象的变量。

#### **SmartArtLayouts.Parent**

返回调用对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SmartArtLayouts** 对象的变量。

适用环境：web

适用平台：windows/linux