**Scenario**



代表工作表上的一个方案。

**说明**

方案是一组经过命名和保存的输入值（被称为 *changing cells*）。**Scenario** 对象是 **Scenarios** 集合的成员。**Scenarios** 集合包含某个工作表的所有定义方案。

**方法**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChangeScenario** | 更改方案以得到一组新的可变单元格及方案值（可选）。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**         | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show**           | 通过在工作表中插入方案的值来显示方案。受影响的单元格为方案的更改单元格。 |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChangingCells** | 返回一个 **Range** 对象，该对象表示方案的可变单元格。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comment**       | 返回或设置一个 **String** 值，它代表与方案相关联的批注。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**        | 返回或设置一个 **Boolean** 值，它指明方案是否已被隐藏。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**         | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**        | 返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**          | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Values**        | 返回一个 **Variant** 数组，它包含方案的可变单元格的当前值。  |

**成员方法**

#### **Scenario.ChangeScenario**

更改方案以得到一组新的可变单元格及方案值（可选）。

**语法**

**express.ChangeScenario(ChangingCells, Values)**

*express*   一个代表 **Scenario** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ChangingCells* | 必选          | **Variant**  | 指定方案中新的可变单元格集合的 Range 对象。这些可变单元格必须与方案位于同一工作表中。 |
| *Values*        | 可选          | **Variant**  | 包含可变单元格的新方案值的数组。如果省略此参数，将假定方案值为这些可变单元格的当前值。 |

**返回值**

Variant

**说明**

如果指定 *Values* 参数，则数组中每一元素必须对应于 *ChangingCells* 区域中的相应单元格；否则，ET 就会产生错误。

**示例**

本示例将方案一的可变单元格设为工作表 Sheet1 上的区域 A1:A10。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Scenarios(1).ChangeScenario(Worksheets.Item("Sheet1").Range("A1:A10"))` |

#### **Scenario.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Scenario** 对象的变量。

**返回值**

Variant

#### **Scenario.Show**

通过在工作表中插入方案的值来显示方案。受影响的单元格为方案的更改单元格。

**语法**

**express.Show()**

*express*   一个代表 **Scenario** 对象的变量。

**返回值**

Variant

**成员属性**

#### **Scenario.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Scenario** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook        if(myObject.Application.Value == "ET") { 　　　　    MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **Scenario.ChangingCells**

返回一个 **Range** 对象，该对象表示方案的可变单元格。只读。

**语法**

**express.ChangingCells**

*express*   一个代表 **Scenario** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定 Sheet1 中方案一的可变单元格。*/ function test(){ 　　　　Worksheets.Item("Sheet1").Activate() 　　　　ActiveSheet.Scenarios(1).ChangingCells.Select() }` |

#### **Scenario.Comment**

返回或设置一个 **String** 值，它代表与方案相关联的批注。

**语法**

**express.Comment**

*express*   一个代表 **Scenario** 对象的变量。

**说明**

批注文本不能超过 255 个字符。

**示例**

此示例设置 Sheet1 中第一个方案的批注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Worksheets.Item("Sheet1").Scenarios(1).Comment ="Worst case July 1993 sales"` |

#### **Scenario.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Scenario** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Scenario.Hidden**

返回或设置一个 **Boolean** 值，它指明方案是否已被隐藏。

**语法**

**express.Hidden**

*express*   一个代表 **Scenario** 对象的变量。

**说明**

此属性的默认设置为 **False**。

请不要将此属性与 **FormulaHidden** 属性混淆。

#### **Scenario.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Scenario** 对象的变量。

#### **Scenario.Locked**

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

*express*   一个代表 **Scenario** 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

#### **Scenario.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Scenario** 对象的变量。

#### **Scenario.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Scenario** 对象的变量。

#### **Scenario.Values**

返回一个 **Variant** 数组，它包含方案的可变单元格的当前值。

**语法**

**express.Values**

*express*   一个代表 **Scenario** 对象的变量。

适用环境：web

适用平台：windows/linux