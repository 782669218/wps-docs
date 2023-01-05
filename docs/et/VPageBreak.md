**VPageBreak**



代表一个垂直分页符。

**说明**

**VPageBreak** 对象是 [**VPageBreaks** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/VPageBreaks/VPageBreaks%20.htm#jsObject_VPageBreaks)集合的成员。

使用 **VPageBreaks**(*index*)（其中 *index* 是该分页符的分页符索引号）可返回一个 **VPageBreak** 对象。下例更改第一个垂直分页符的位置。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")` |

**方法**

|                                                              | 名称        | 说明                       |
| ------------------------------------------------------------ | ----------- | -------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**  | 删除对象。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DragOff** | 将一个分页符拖出打印区域。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Extent**      | 返回指定分页符的类型：全屏或仅在打印区域内。可为以下任一 **XlPageBreakExtent** [ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPageBreakExtent%20%E6%9E%9A%E4%B8%BE.html)常量：**xlPageBreakFull**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPageBreakExtent%20%E6%9E%9A%E4%B8%BE.html)或 **xlPageBreakPartial**。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**    | 返回或设置定义分页符位置的单元格（**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **VPageBreak.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **VPageBreak** 对象的变量。

#### **VPageBreak.DragOff**

将一个分页符拖出打印区域。

**语法**

**express.DragOff(Direction, RegionIndex)**

*express*   一个代表 **VPageBreak** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Direction*   | 必选          | **XlDirection** | 分页符拖动方向。                                             |
| *RegionIndex* | 必选          | **Long**        | 分页符的打印区域索引（当用户按下鼠标按钮拖动分页符时鼠标指针所在的位置）。如果打印区域是连续的，则只有一个打印区域。如果打印区域不是连续的，则有多个打印区域。 |

**说明**

该方法主要用于宏记录器。使用 **Delete** 方法可在 JS 中删除分页符。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将活动工作表的第一个垂直分页符拖出第一个打印区域的右边界，以删除该分页符*/ Application.ActiveSheet.VPageBreaks.Item(1).DragOff(xlToRight, 1)` |

**成员属性**

#### **VPageBreak.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **VPageBreak** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **VPageBreak.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **VPageBreak** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **VPageBreak.Extent**

返回指定分页符的类型：全屏或仅在打印区域内。可为以下任一 **XlPageBreakExtent** [ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPageBreakExtent%20%E6%9E%9A%E4%B8%BE.html)常量：**xlPageBreakFull**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlPageBreakExtent%20%E6%9E%9A%E4%B8%BE.html)或 **xlPageBreakPartial**。**Long** 类型，只读。

**语法**

**express.Extent**

*express*   一个代表 **VPageBreak** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示全屏水平分页符和打印区水平分页符的总数*/ function test() {     let cFull = 0     let cPartial = 0     for(let i = 1; i <= Application.Worksheets.Item(1).HPageBreaks.Count; i++){         if(Application.Worksheets.Item(1).HPageBreaks.Item(i).Extent == xlPageBreakFull){             cFull = cFull + 1         }         else{             cPartial = cPartial + 1         }     }     alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks") }` |

#### **VPageBreak.Location**

返回或设置定义分页符位置的单元格（**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象）。水平分页符与定位单元格的上边缘对齐；垂直分页符与定位单元格的左边缘对齐。**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)类型，可读写。

**语法**

**express.Location**

*express*   一个代表 **VPageBreak** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例移动垂直分页符的位置*/ Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")` |

#### **VPageBreak.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **VPageBreak** 对象的变量。

#### **VPageBreak.Type**

返回指定对象的父对象。只读。

**语法**

**express.Type**

*express*   一个代表 **VPageBreak** 对象的变量。

适用环境：web

适用平台：windows/linux