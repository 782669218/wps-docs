**WorksheetFunction**

（由于原官方文档问题，排版不是很正常，建议到[官网](https://open.wps.cn/docs/client/wpsLoad)看）

用作可从 Visual Basic 中调用的 ET 工作表函数的容器。

**说明**

使用 **WorksheetFunction** 属性可返回 **WorksheetFunction** 对象。下例显示给区域 A1:A10 应用 **Min** 工作表函数的结果。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10") let answer = Application.WorksheetFunction.Min(myRange) alert(answer) }` |

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AccrInt**         | 返回定期付息有价证券的应计利息。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AccrIntM**        | 返回到期一次性付息有价证券的应计利息。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Acos**            | 返回数字的反余弦值。反余弦值是余弦值为 *Arg1* 的角度。返回的角度以弧度给出，范围从 0（零）到 pi。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Acosh**           | 返回数字的反双曲余弦值。数字必须大于或等于 1。反双曲余弦值的双曲余弦值为 *Arg1*，因此 Acosh(Cosh(number)) 等于 *Arg1*。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Aggregate**       | 返回列表或数据库中的一个聚合。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AmorDegrc**       | 返回每个结算期的折旧值。此函数为法国会计系统提供。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AmorLinc**        | 返回每个结算期的折旧值。此函数为法国会计系统提供。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **And**             | 如果其所有参数都为 TRUE，则返回 TRUE；如果一个或多个参数为 FALSE，则返回 FALSE。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Asc**             | 对于双字节字符集 (DBCS) 语言，将全角（双字节）字符更改为半角（单字节）字符。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Asin**            | 返回一个数字的反正弦值。反正弦值是一个角度，这个角度的正弦值为 *Arg1*。返回的角度采用弧度的形式，其范围从 -pi/2 到 pi/2。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Asinh**           | 返回数字的反双曲正弦值。反双曲正弦值的双曲正弦值为 *Arg1*，因此 Asinh(Sinh(number)) 等于 *Arg1*。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Atan2**           | 返回指定的 x 和 y 坐标值的反正切值。反正切值是角度，是从 x 轴到通过原点 (0, 0) 和坐标点 (x_num, y_num) 的直线之间的夹角。该角度用弧度给出，介于 -pi 和 pi 之间（不包括 -pi）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Atanh**           | 返回数字的反双曲正切值。数字必须介于 -1 和 1 之间（不包括 -1 和 1）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AveDev**          | 返回多个数据点与其平均值的绝对偏差的平均值。AveDev 是对数据集中可变性的度量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Average**         | 返回参数的平均值（算术平均值）。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AverageIf**       | 返回区域内满足给定条件的所有单元格的平均值（算术平均值）。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AverageIfs**      | 返回满足多个条件的所有单元格的平均值（算术平均值）。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BahtText**        | 将数字转换为泰语文本并添加后缀“泰铢”。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BesselI**         | 返回修正的 Bessel 函数，它等效于计算纯虚数参数值的 Bessel 函数。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BesselJ**         | 返回 Bessel 函数。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BesselK**         | 返回修正 Bessel 函数，它等效于根据纯虚数参数计算的 Bessel 函数。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BesselY**         | 返回 Bessel 函数，该函数也称为 Weber 函数或 Neumann 函数。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Beta_Dist**       | 返回 Beta 累积分布函数。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Beta_Inv**        | 返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = Beta_Dist(x,...)，则 Beta_Inv(probability,...) = x。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BetaDist**        | 返回 Beta 累积分布函数。 **要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。有关新函数的详细信息，请参阅 Beta_Dist 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BetaInv**         | 返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = BetaDist(x,...)，则 BetaInv(probability,...) = x。 **要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。有关新函数的详细信息，请参阅 Beta_Inv 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Bin2Dec**         | 将二进制数转换为十进制数。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Bin2Hex**         | 将二进制数转换为十六进制数。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Bin2Oct**         | 将二进制数转换为八进制数。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Binom_Dist**      | 返回一元二项式分布的概率。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Binom_Inv**       | 返回一元二项式概率分布的反函数。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BinomDist**       | 返回一元二项式分布的概率。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ceiling**         | 返回向上舍入（远离零）到最接近的 significance 的倍数的 number。 **要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。有关新函数的详细信息，请参阅 **Ceiling_Precise** 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ceiling_Precise** | 返回舍入到最接近的有效位倍数的指定数字。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChiDist**         | 返回 χ2 分布的单尾概率。此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。有关新函数的详细信息，请参阅 ChiSq_Dist_RT 和 ChiSq_Dist 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChiInv**          | 返回 χ2 分布的单尾概率的反函数。此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。有关新函数的详细信息，请参阅 ChiSq_Inv_RT 和 ChiSq_Inv 方法。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChiSq_Dist**      | 返回 χ2 分布。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChiSq_Dist_RT**   | 返回 χ2 分布的右尾概率。                                     |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象。可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **WorksheetFunction.AccrInt**

返回定期付息有价证券的应计利息。

**语法**

**express.AccrInt(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                            |
| -------- | ------------- | ------------ | ----------------------------------- |
| *Arg1*   | 必选          | **Variant**  | Issue date - 债券的发行日期。       |
| *Arg2*   | 必选          | **Variant**  | First interest - 债券的首次计息日。 |
| *Arg3*   | 必选          | **Variant**  | Settlement - 债券的结算日。         |
| *Arg4*   | 必选          | **Variant**  | Rate - 债券的年息票利率。           |
| *Arg5*   | 必选          | **Variant**  | Par - 债券的票面价值。              |
| *Arg6*   | 必选          | **Variant**  | Frequency - 每年支付息票的次数。    |
| *Arg7*   | 可选          | **Variant**  | Basis - 要使用的日计数基准类型。    |

**返回值**

Double

**说明**

日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

下表描述了可用于 *Arg5* 的值。

| 基准     | 日计数基准                       |
| -------- | -------------------------------- |
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1        | 实际天数/实际天数                |
| 2        | 实际天数/360                     |
| 3        | 实际天数/365                     |
| 4        | 欧洲 30/360                      |

#### **WorksheetFunction.AccrIntM**

返回到期一次性付息有价证券的应计利息。

**语法**

**express.AccrIntM(Arg1, Arg2, Arg3, Arg4, Arg5)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                              |
| -------- | ------------- | ------------ | ----------------------------------------------------- |
| *Arg1*   | 必选          | **Variant**  | 债券的发行日。                                        |
| *Arg2*   | 必选          | **Variant**  | 债券的到期日。                                        |
| *Arg3*   | 必选          | **Variant**  | 债券的年息票利率。                                    |
| *Arg4*   | 必选          | **Variant**  | 债券的票面价值。如果省略 par，ACCRINTM 则使用 $1000。 |
| *Arg5*   | 可选          | **Variant**  | 要使用的日计数基准类型。                              |

**返回值**

Double

**说明**

**要点：**日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

下表描述了可用于 *Arg5* 的值。

| 基准     | 日计数基准                       |
| -------- | -------------------------------- |
| 0 或省略 | 美国（美国证券交易商协会）30/360 |
| 1        | 实际天数/实际天数                |
| 2        | 实际天数/360                     |
| 3        | 实际天数/365                     |
| 4        | 欧洲 30/360                      |

以下列表包含在使用 ACCRINTM 时要注意的信息
ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日是序列号 1，2008 年 1 月 1 日是序列号 39,448，这是因为它距 1900 年 1 月 1 日有 39,448 天。
Issue、maturity 和 basis 将被截尾取整。
如果 issue 或 maturity 是无效的日期，则 ACCRINTM 将生成一个错误。
如果 rate ≤ 0 或 par ≤ 0，则 ACCRINTM 将产生一个错误。
如果 basis < 0 或 basis > 4，则 ACCRINTM 将产生一个错误。
如果 issue ≥ maturity，则 ACCRINTM 将产生一个错误。
ACCRINTM 的计算公式如下：

其中：

A = 按月计算的应计天数。在计算到期付息的利息时指发行日与到期日之间的天数。

D = 年基准数。



#### **WorksheetFunction.Acos**

返回数字的反余弦值。反余弦值是余弦值为 *Arg1* 的角度。返回的角度以弧度给出，范围从 0（零）到 pi。

**语法**

**express.Acos(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                  |
| -------- | ------------- | ------------ | ----------------------------------------- |
| *Arg1*   | 必选          | **Double**   | 所需角度的余弦值，必须介于 -1 到 1 之间。 |

**说明**

如果要将结果从弧度转换为度，则请将结果乘以 180/PI()，或者使用 Degrees 方法。

#### **WorksheetFunction.Acosh**

返回数字的反双曲余弦值。数字必须大于或等于 1。反双曲余弦值的双曲余弦值为 *Arg1*，因此 Acosh(Cosh(number)) 等于 *Arg1*。

**语法**

**express.Acosh(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                  |
| -------- | ------------- | ------------ | ------------------------- |
| *Arg1*   | 必选          | **Double**   | 等于或大于 1 的任意实数。 |

**返回值**

Double

#### **WorksheetFunction.Aggregate**

返回列表或数据库中的一个聚合。

**语法**

**express.Aggregate(Arg1, Arg2, Arg3, Arg4 - Arg 30)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*          | 必选          | **Double**   | Function_num - 一个介于 1 到 19 之间的数字，用于指定要使用的函数。 Function_num 函数 1 AVERAGE 2 COUNT 3 COUNTA 4 MAX 5 MIN 6 PRODUCT 7 STDEV.S 8 STDEV.P 9 SUM 10 VAR.S 11 VAR.P 12 MEDIAN 13 MODE.SNGL 14 LARGE 15 SMALL 16 PERCENTILE.INC 17 QUARTILE.INC 18 PERCENTILE.EXC 19 QUARTILE.EXC Arg2 必选 Double 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值 |
| *Arg2*          | 必选          | **Double**   | 选项 - 一个数值，用于确定要在函数的求值范围中忽略的值。 选项 行为 0 或省略 忽略嵌套的 SUBTOTAL 和 AGGREGATE 函数 1 忽略隐藏行以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 2 忽略错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 3 忽略隐藏行、错误值以及嵌套的 SUBTOTAL 和 AGGREGATE 函数 4 忽略空值 5 忽略隐藏行 6 忽略错误值 7 忽略隐藏行和错误值 |
| *Arg3*          | 必选          | **Range**    | Ref1 - 带有多个要对其计算聚合值的数值参数的函数的第一个数值参数。 |
| *Arg4 - Arg 30* | 可选          | **Variant**  | Ref2 - Ref30 - 要对其计算聚合值的 2 到 30 个数值参数。       |

**返回值**

Double

**说明**

以下约束将会根据 Function_num 值应用于 Ref 参数 (Arg3 - Arg 30)。
Function_num	Ref1	Ref2	Ref3、Ref4…
1-13	有效类型：
任何单元格引用
并集
交集
定义的名称
结构化引用

无效类型：
实际数据
数组
有效类型：
任何单元格引用
并集
交集
定义的名称
结构化引用

无效类型：
实际数据
数组
有效类型：
任何单元格引用
并集
交集
定义的名称
结构化引用

无效类型：
实际数据
数组
14-17	有效类型：
任何单元格引用
并集
交集
定义的名称
结构化引用
实际数据
数组
有效类型：
任何单元格引用
并集
交集
定义的名称
结构化引用
实际数据
数组
不允许使用引用
如果需要但未提供第二个 Ref 参数，AGGREGATE 将会返回错误 #VALUE!。
如果有一个或多个引用为三维引用，AGGREGATE 将会返回错误值 #VALUE!。



#### **WorksheetFunction.AmorDegrc**

返回每个结算期的折旧值。此函数为法国会计系统提供。

**语法**

**express.AmorDegrc(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Arg1*   | 必选          | **Variant**  | 资产原值。               |
| *Arg2*   | 必选          | **Variant**  | 资产的购买日。           |
| *Arg3*   | 必选          | **Variant**  | 第一个结算期的结束日期。 |
| *Arg4*   | 必选          | **Variant**  | 资产在寿命结束时的残值。 |
| *Arg5*   | 必选          | **Variant**  | 结算期。                 |
| *Arg6*   | 必选          | **Variant**  | 折旧率。                 |
| *Arg7*   | 可选          | **Variant**  | 要使用的年基准数。       |

**说明**

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。该方法与 AmorLinc 相似，不同之处在于该方法中用于计算的折旧系数取决于资产的寿命。

下表描述了 *Arg7* 中使用的值。

| Basis    | 日期系统                |
| -------- | ----------------------- |
| 0 或省略 | 360 天（NASD 方法）     |
| 1        | 实际天数                |
| 3        | 一年 365 天             |
| 4        | 一年 360 天（欧洲方法） |

 ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。ET for the Macintosh 使用另外一个默认日期系统。
此函数将返回折旧值，截止到资产寿命的最后一个期间，或直到累积折旧值大于资产原值减去残值后所得的结果。
折旧系数为：
资产的寿命 (1/rate)	折旧系数
3 到 4 年	1.5
5 到 6 年	2
6 年以上	2.5
最后一个期间之前的那个期间的折旧率将增加到 50%，最后一个期间的折旧率将增加到 100%。
如果资产的寿命在 0（零）到 1、1 到 2、2 到 3 或 4 到 5 之间，则将返回错误值 #NUM!。



#### **WorksheetFunction.AmorLinc**

返回每个结算期的折旧值。此函数为法国会计系统提供。

**语法**

**express.AmorLinc(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Arg1*   | 必选          | **Variant**  | 资产原值。               |
| *Arg2*   | 必选          | **Variant**  | 资产的购买日。           |
| *Arg3*   | 必选          | **Variant**  | 第一个结算期的结束日期。 |
| *Arg4*   | 必选          | **Variant**  | 资产在寿命结束时的残值。 |
| *Arg5*   | 必选          | **Variant**  | 结算期。                 |
| *Arg6*   | 必选          | **Variant**  | 折旧率。                 |
| *Arg7*   | 可选          | **Variant**  | 要使用的年基准数。       |

**返回值**

Double

**说明**

如果某项资产是在结算期的中期购入的，则应考虑按比例折旧。

下表描述了 *Arg7* 所使用的值。

| Basis    | 日期系统                |
| -------- | ----------------------- |
| 0 或省略 | 360 天（NASD 方法）     |
| 1        | 实际天数                |
| 3        | 一年 365 天             |
| 4        | 一年 360 天（欧洲方法） |

**要点**??日期应使用 DATE 函数输入，或者作为其他公式或函数的结果输入。例如，使用 DATE(2008,5,23) 输入 2008 年 5 月 23 日。如果日期以文本形式输入，将会出现问题。

ET 以序数形式存储日期以使其可用于计算。默认情况下，1900 年 1 月 1 日的序数是 1；2008 年 1 月 1 日的序数是 39448，因为该日期距 1900 年 1 月 1 日有 39,448 天。ET for the Macintosh 使用另外一个默认日期系统。



#### **WorksheetFunction.And**

如果其所有参数都为 TRUE，则返回 TRUE；如果一个或多个参数为 FALSE，则返回 FALSE。

**语法**

**express.And(Arg1 - Arg30)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                           |
| -------------- | ------------- | ------------ | -------------------------------------------------- |
| *Arg1 - Arg30* | 必选          | **Variant**  | 1 到 30 个需要进行检验的条件，可为 TRUE 或 FALSE。 |

**返回值**

Boolean

**说明**

- 参数的计算结果必须为逻辑值（如 TRUE 或 FALSE），或者参数必须是包含逻辑值的数组或引用。
- 如果数组或引用参数中包含文本或空单元格，则这些值将被忽略。
- 如果指定区域内不包含逻辑值，则此方法将生成一个错误值。



#### **WorksheetFunction.Asc**

对于双字节字符集 (DBCS) 语言，将全角（双字节）字符更改为半角（单字节）字符。

**语法**

**express.Asc(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **String**   | 文本或对包含要更改的文本的单元格的引用。如果文本中不包含任何全角字母，则文本不会更改。 |

**返回值**

String

#### **WorksheetFunction.Asin**

返回一个数字的反正弦值。反正弦值是一个角度，这个角度的正弦值为 *Arg1*。返回的角度采用弧度的形式，其范围从 -pi/2 到 pi/2。

**语法**

**express.Asin(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                  |
| -------- | ------------- | ------------ | ----------------------------------------- |
| *Arg1*   | 必选          | **Double**   | 所需角度的正弦值，必须介于 -1 到 1 之间。 |

**返回值**

Double

**说明**

要用度表示反正弦值，请将结果乘以 180/PI( ) 或使用 Degrees 方法。

#### **WorksheetFunction.Asinh**

返回数字的反双曲正弦值。反双曲正弦值的双曲正弦值为 *Arg1*，因此 Asinh(Sinh(number)) 等于 *Arg1*。

**语法**

**express.Asinh(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**   |
| -------- | ------------- | ------------ | ---------- |
| *Arg1*   | 必选          | **Double**   | 任意实数。 |

**返回值**

Double

#### **WorksheetFunction.Atan2**

返回指定的 x 和 y 坐标值的反正切值。反正切值是角度，是从 x 轴到通过原点 (0, 0) 和坐标点 (x_num, y_num) 的直线之间的夹角。该角度用弧度给出，介于 -pi 和 pi 之间（不包括 -pi）。

**语法**

**express.Atan2(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**      |
| -------- | ------------- | ------------ | ------------- |
| *Arg1*   | 必选          | **Double**   | 点的 x 坐标。 |
| *Arg2*   | 必选          | **Double**   | 点的 y 坐标。 |

**返回值**

Double

**说明**

- 结果为正表示从 x 轴逆时针旋转的角度；结果为负表示从 x 轴顺时针旋转的角度。
- Atan2(a,b) 等于 Atan(b/a)，不同之处在于 Atan2 中的 a 可等于 0。
- 如果 x_num 和 y_num 为 0，则 Atan2 将返回一个错误值。
- 要用度表示反正切值，请将结果乘以 180/PI( ) 或使用 Degrees 方法。

#### **WorksheetFunction.Atanh**

返回数字的反双曲正切值。数字必须介于 -1 和 1 之间（不包括 -1 和 1）。

**语法**

**express.Atanh(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                      |
| -------- | ------------- | ------------ | ----------------------------- |
| *Arg1*   | 必选          | **Double**   | 介于 1 和 -1 之间的任意实数。 |

**返回值**

Double

**说明**

反双曲正切值的双曲正切值为 *Arg1*，因此 Atanh(Tanh(number)) 等于 *Arg1*。

#### **WorksheetFunction.AveDev**

返回多个数据点与其平均值的绝对偏差的平均值。AveDev 是对数据集中可变性的度量。

**语法**

**express.AveDev(Arg1 - Arg 30)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1 - Arg 30* | 必选          | **Variant**  | 要计算其绝对偏差平均值的 1 到 30 个参数。也可以不使用这种用逗号分隔参数的形式，而使用一个数组或一个对数组的引用。 |

**返回值**

Double

**说明**

- 
- AveDev 受输入数据中的度量单位的影响。
- 参数要么是数字，要么是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 平均偏差的计算公式为：

 **WorksheetFunction.Average**

返回参数的平均值（算术平均值）。

**语法**

**express.Average(Arg1 - Arg30)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                            |
| -------------- | ------------- | ------------ | ----------------------------------- |
| *Arg1 - Arg30* | 必选          | **Variant**  | 要求其平均值的 1 到 30 个数值参数。 |

**返回值**

Double

**说明**

- 参数可以是数字，也可以是包含数字的名称、数组或引用。
- 直接键入参数列表的逻辑值和数字的文本表示也包括在内。
- 如果数组或引用参数包含文本、逻辑值或空单元格，则这些值将被忽略；但含有零值的单元格包括在内。
- 如果参数为错误值或不能转换为数字的文本，则将导致错误。
- 如果要对引用中的逻辑值和数字的文本表示进行计算，请使用 AVERAGEA 函数。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| Average 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

**提示** 对单元格求平均值时，应牢记空单元格与含零值单元格之间的区别，尤其在已经清除了**“视图”**选项卡上的**“零值”**复选框（**“工具”**菜单上的**“选项”**命令）的情况下更应如此。空单元格不计算在内，但会计算含零值的单元格。



#### **WorksheetFunction.AverageIf**

返回区域内满足给定条件的所有单元格的平均值（算术平均值）。

**语法**

**express.AverageIf(Arg1, Arg2, Arg3)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Range**    | 要求其平均值的一个或多个单元格。                             |
| *Arg2*   | 必选          | **Variant**  | 定义将对哪些单元格求平均值的条件，其形式可以为数字、表达式、单元格引用或文本。例如，条件可以表示为 32、"32"、">32"、"apples" 或 B4。 |
| *Arg3*   | 可选          | **Variant**  | 要求其平均值的实际单元格集合。如果省略，则使用 range。       |

**返回值**

Double

**说明**

- range 内包含 TRUE 或 FALSE 的单元格将被忽略。

- 如果 range 或 average_range 中的某个单元格为空单元格，则 AverageIf 会将其忽略。

- 如果条件中的单元格为空，则 AverageIf 会将其作为 0 值处理。

- 如果区域中没有满足条件的单元格，则 AverageIf 将生成一个错误值。

- 可以在条件中使用通配符，包括问号 (?) 和星号 (*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

- Average_range 的大小和形状不必与 range 相同。求其平均值的实际单元格的确定方法如下：使用 average_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如：

- 如果 range 为

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| AverageIf 方法衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

#### **WorksheetFunction.AverageIfs**

返回满足多个条件的所有单元格的平均值（算术平均值）。

**语法**

**express.AverageIfs(Arg1 - Arg30)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                             |
| -------------- | ------------- | ------------ | ------------------------------------ |
| *Arg1 - Arg30* | 必选          | **Range**    | 在其中计算相关条件的一个或多个区域。 |

**返回值**

Double

**说明**

- 如果 average_range 中的某个单元格为空单元格，则 AverageIfs 会将其忽略。
- 如果条件区域中的某个单元格为空，则 AverageIfs 会将其作为 0 值处理。
- range 中包含 TRUE 的单元格作为 1 计算；range 中包含 FALSE 的单元格作为 0（零）计算。
- 仅在每个单元格中指定的对应条件都为 True 时，才会在平均值计算过程中使用 average_range 中的该单元格。
- 如果 average_range 中的单元格均为空，或包含无法转换为数字的文本值，则 AverageIfs 将生成一个错误。
- 如果没有满足所有条件的单元格，则 AverageIfs 将生成一个错误值。
- 可以在条件中使用通配符，包括问号 (?) 和星号 (*)。问号可匹配任意的单个字符；星号可匹配任意一串字符。如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。
- 每个 criteria_range 的大小和形状不必与 average_range 相同。计算其平均值的实际单元格的确定方法如下：使用 criteria_range 中左上角的单元格作为起始单元格，然后将与 range 的大小和形状对应的所有单元格包含到其中。例如：

| 如果 average_range 为 | 并且 criteria_range 为 | 则实际纳入计算的单元格为 |
| --------------------- | ---------------------- | ------------------------ |
| A1:A5                 | B1:B5                  | B1:B5                    |
| A1:A5                 | B1:B3                  | B1:B5                    |
| A1:B4                 | C1:D4                  | C1:D4                    |
| A1:B4                 | C1:C2                  | C1:D4                    |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| AverageIfs 函数衡量趋中性，趋中性是统计分布中一组数字的中心位置。三种最常见的趋中性衡量方式为： |

- **平均值** 即算术平均值，是通过将一组数字相加，然后除以数字个数得到的。例如，2、3、3、5、7 和 10 的平均值为 30 除以 6，即为 5。
- **中值** 是一组数字中居于中间的数，即在这组数字中，有一半的数字的值比它大，有一半的数字的值比它小。例如，2、3、3、5、7 和 10 的中值为 4。
- **众值** 即在一组数字中出现次数最多的数字。例如，2、3、3、5、7 和 10 的众值为 3。

对于对称分布的一组数字，这三种趋中性衡量方式完全相同。对于偏态分布的一组数字，这些衡量方式可能会不同。

#### **WorksheetFunction.BahtText**

将数字转换为泰语文本并添加后缀“泰铢”。

**语法**

**express.BahtText(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Double**   | 要转换成文本的数字、对包含数字的单元格的引用或计算结果为数字的公式。 |

**返回值**

String

**说明**

在 ET for Windows 中，可以使用**“控制面板”**中的“区域设置”或**“区域选项”**将泰铢格式更改为其他样式。

在 ET for the Macintosh 中，可以使用**“数字控制面板”**将泰铢数字格式更改为其他样式。

#### **WorksheetFunction.BesselI**

返回修正的 Bessel 函数，它等效于计算纯虚数参数值的 Bessel 函数。

**语法**

**express.BesselI(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                           |
| -------- | ------------- | ------------ | -------------------------------------------------- |
| *Arg1*   | 必选          | **Variant**  | 要计算的函数值。                                   |
| *Arg2*   | 必选          | **Variant**  | Bessel 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 
- 如果 x 是非数值，BesselI 将返回错误值 #VALUE!。
- 如果 n 是非数值，BesselI 将生成一个错误值。
- 如果 n < 0，BesselI 将生成一个错误值。
- 变量 x 的 n 阶修正 Bessel 函数为：

 **WorksheetFunction.BesselJ**

返回 Bessel 函数。

**语法**

**express.BesselJ(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                           |
| -------- | ------------- | ------------ | -------------------------------------------------- |
| *Arg1*   | 必选          | **Variant**  | 计算的函数值。                                     |
| *Arg2*   | 必选          | **Variant**  | Bessel 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 
- 如果 x 为非数值型，则 BesselJ 将生成一个错误值。
- 如果 n 为非数值型，则 BesselJ 将返回一个错误值。
- 如果 n < 0，则 BesselJ 将生成一个错误值。
- 变量 x 的 n 阶 Bessel 函数为：

 其中：

![img]()

为 γ 函数。

#### **WorksheetFunction.BesselK**

返回修正 Bessel 函数，它等效于根据纯虚数参数计算的 Bessel 函数。

**语法**

**express.BesselK(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                    |
| -------- | ------------- | ------------ | ------------------------------------------- |
| *Arg1*   | 必选          | **Variant**  | 计算的函数值。                              |
| *Arg2*   | 必选          | **Variant**  | 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 
- 如果 x 为非数值型，则 BesselK 将生成一个错误值。
- 如果 n 为非数值型，则 BesselK 将生成一个错误值。
- 如果 n < 0，则 BesselK 将生成一个错误值。
- 变量 x 的 n 阶修正 Bessel 函数为：

 其中 Jn 和 Yn 分别为 J 和 Y 的 Bessel 函数。 

#### **WorksheetFunction.BesselY**

返回 Bessel 函数，该函数也称为 Weber 函数或 Neumann 函数。

**语法**

**express.BesselY(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                    |
| -------- | ------------- | ------------ | ------------------------------------------- |
| *Arg1*   | 必选          | **Variant**  | 计算的函数值。                              |
| *Arg2*   | 必选          | **Variant**  | 函数的阶。如果 n 不是整数，则将被截尾取整。 |

**返回值**

Double

**说明**

- 
- 如果 x 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n 为非数值型，则 BesselY 将生成一个错误值。
- 如果 n < 0，则 BesselY 将生成一个错误值。
- 变量 x 的 n 阶 Bessel 函数为：

 **WorksheetFunction.Beta_Dist**

返回 Beta 累积分布函数。

**语法**

**express.Beta_Dist(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Double**   | 介于 A 和 B 之间的值，用来计算函数的值。                     |
| *Arg2*   | 必选          | **Double**   | 该分布的 Alpha 参数。                                        |
| *Arg3*   | 必选          | **Double**   | 该分布的 Beta 参数。                                         |
| *Arg4*   | 可选          | **Variant**  | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 BETA.DIST 将返回累积分布函数；如果为 FALSE，则将返回概率密度函数。 |
| *Arg5*   | 可选          | **Variant**  | x 所属区间的可选下界。                                       |
| *Arg6*   | 可选          | **Variant**  | x 所属区间的可选上界。                                       |

**返回值**

Double

**说明**

Beta 分布通常用于研究样本中某些事物变化的百分比，例如人们一天中用来看电视的时间所占的比例：

- 如果任一参数为非数值型，则 Beta_Dist 将返回错误值 #VALUE!。
- 如果 Alpha ≤ 0 或 Beta ≤ 0，则 Beta_Dist 将生成一个错误值。
- 如果 x < A、x > B 或 A = B，则 Beta_Dist 将生成一个错误值。
- 如果忽略 A 和 B 的值（即忽略下界和上界），则 Beta_Dist 将使用标准累积 Beta 分布，使 A = 0 且 B = 1。

#### **WorksheetFunction.Beta_Inv**

返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = Beta_Dist(x,...)，则 Beta_Inv(probability,...) = x。

**语法**

**express.Beta_Inv(Arg1, Arg2, Arg3, Arg4, Arg5)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Arg1*   | 必选          | **Double**   | 与 Beta 分布相关的概率。 |
| *Arg2*   | 必选          | **Double**   | 该分布的 Alpha 参数。    |
| *Arg3*   | 必选          | **Double**   | 该分布的 Beta 参数。     |
| *Arg4*   | 可选          | **Variant**  | x 所属区间的可选下界。   |
| *Arg5*   | 可选          | **Variant**  | x 所属区间的可选上界。   |

**返回值**

Double

**说明**

在项目计划中，如果给定预期完成时间和变化率，则可使用 Beta 分布来建立可能完成时间的模型：

- 如果任一参数为非数值型，则 Beta_Inv 将生成一个错误值。
- 如果 Alpha ≤ 0 或 Beta ≤ 0，则 Beta_Inv 将生成一个错误值。
- 如果 probability ≤ 0 或 probability > 1，则 Beta_Inv 将生成一个错误值。
- 如果忽略 A 和 B 的值（即忽略下界和上界），则 Beta_Inv 将使用标准累积 Beta 分布，使 A = 0 且 B = 1。

如果给定 probability 的值，Beta_Inv 将使用 Beta_Dist(x, alpha, beta, TRUE, A, B) = probability 求解值 x。因此，Beta_Inv 的精度取决于 Beta_Dist 的精度。Beta_Inv 使用迭代搜索技术。 

#### **WorksheetFunction.BetaDist**

返回 Beta 累积分布函数。

**要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 Beta_Dist 方法。

**语法**

**express.BetaDist(Arg1, Arg2, Arg3, Arg4, Arg5)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                 |
| -------- | ------------- | ------------ | ---------------------------------------- |
| *Arg1*   | 必选          | **Double**   | 介于 A 和 B 之间的值，用来计算函数的值。 |
| *Arg2*   | 必选          | **Double**   | 分布参数。                               |
| *Arg3*   | 必选          | **Double**   | 分布参数。                               |
| *Arg4*   | 可选          | **Variant**  | x 所属区间的可选下界。                   |
| *Arg5*   | 可选          | **Variant**  | x 所属区间的可选上界。                   |

**返回值**

Double

**说明**

Beta 分布通常用于研究样本中某些事物变化的百分比，例如人们一天中用来看电视的时间所占的比例。

- 如果任一参数为非数值型，则 BetaDist 将返回错误值 #VALUE!。
- 如果 alpha ≤ 0 或 beta ≤ 0，则 BetaDist 将生成一个错误值。
- 如果 x < A、x > B 或 A = B，则 BetaDist 将生成一个错误值。
- 如果省略 A 和 B 的值，则 BetaDist 将使用标准累积 Beta 分布，使 A = 0，B = 1。

#### **WorksheetFunction.BetaInv**

返回指定的 Beta 分布的累积分布函数的反函数。即，如果 probability = BetaDist(x,...)，则 BetaInv(probability,...) = x。

**要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 Beta_Inv 方法。

**语法**

**express.BetaInv(Arg1, Arg2, Arg3, Arg4, Arg5)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Arg1*   | 必选          | **Double**   | 与 Beta 分布相关的概率。 |
| *Arg2*   | 必选          | **Double**   | 该分布的 Alpha 参数。    |
| *Arg3*   | 必选          | **Double**   | 该分布的 Beta 参数。     |
| *Arg4*   | 可选          | **Variant**  | x 所属区间的可选下界。   |
| *Arg5*   | 可选          | **Variant**  | x 所属区间的可选上界。   |

**返回值**

Double

**说明**

在项目计划中，如果已知预期完成时间和变化率，可以使用 beta 分布来建立可能完成时间的模型。

- 如果任一参数是非数值的，BetaInv 就会生成一个错误值。
- 如果 alpha ≤ 0 或 beta ≤ 0，BetaInv 将生成一个错误值。
- 如果 probability ≤ 0 或 probability > 1，BetaInv 将生成一个错误值。
- 如果省略 A 或 B 的值，BetaInv 将使用标准累积 beta 分布，因此，A = 0，B = 1。

如果已知 probability 的值，BetaInv 将求出 x 值以使等式 BetaDist(x, alpha, beta, A, B) = probability 成立。因此，BetaInv 的精度取决于 BetaDist 的精度。BetaInv 使用迭代搜索技术。如果搜索在 100 次迭代之后尚未收敛，该函数将生成一个错误值。 

#### **WorksheetFunction.Bin2Dec**

将二进制数转换为十进制数。

**语法**

**express.Bin2Dec(Arg1)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Variant**  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |

**返回值**

String

**说明**

如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Dec 将生成一个错误值。

#### **WorksheetFunction.Bin2Hex**

将二进制数转换为十六进制数。

**语法**

**express.Bin2Hex(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Variant**  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |
| *Arg2*   | 可选          | **Variant**  | 要使用的字符数。如果省略 places，Bin2Hex 将用能表示此数的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

**返回值**

String

**说明**

- 如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Hex 将生成一个错误。
- 如果 number 为负数，则 Bin2Hex 将忽略 places，并返回一个以 10 个字符表示的十六进制数。
- 如果 Bin2Hex 所需的字符数比 places 指定的字符数多，则将生成一个错误。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 Bin2Hex 将生成一个错误。
- 如果 places 为负数，则 Bin2Hex 将生成一个错误。
- 

#### **WorksheetFunction.Bin2Oct**

将二进制数转换为八进制数。

**语法**

**express.Bin2Oct(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Variant**  | 要转换的二进制数。Number 不能多于 10 个字符（10 位）。最高位为符号位，其余 9 位为数字位。负数用二进制数的补码表示。 |
| *Arg2*   | 可选          | **Variant**  | 要使用的字符数。如果省略 places，Bin2Oct 将用能表示此数的最少字符来表示。当需要为返回的值填充前导 0（零）时，places 尤其有用。 |

**返回值**

String

**说明**

- 如果 number 不是有效的二进制数或多于 10 个字符（10 位），则 Bin2Oct 将生成一个错误。
- 如果 number 为负数，则 Bin2Oct 将忽略 places，并返回一个以 10 个字符表示的八进制数。
- 如果 Bin2Oct 所需的字符数比 places 指定的字符数多，则将生成一个错误。
- 如果 places 不是整数，则将被截尾取整。
- 如果 places 为非数值型，则 Bin2Oct 将生成一个错误。
- 如果 places 为负数，则 Bin2Oct 将生成一个错误。



#### **WorksheetFunction.Binom_Dist**

返回一元二项式分布的概率。

**语法**

**express.Binom_Dist(Arg1, Arg2, Arg3, Arg4)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Double**   | Number_s - 试验成功的次数。                                  |
| *Arg2*   | 必选          | **Double**   | Trials - 独立试验的次数。                                    |
| *Arg3*   | 必选          | **Double**   | Probability_s - 每次试验成功的概率。                         |
| *Arg4*   | 必选          | **Boolean**  | Cumulative - 一个逻辑值，用于确定函数的形式。如果 cumulative 为 True，则 Binom_Dist 方法返回累积分布函数，即最多存在 number_s 次成功的概率；如果为 False，则返回概率密度函数，即存在 number_s 次成功的概率。 |

**返回值**

Double

**说明**

当任何试验的结果仅包含成功或失败两种情况，当试验是独立试验，以及当整个试验过程中成功的概率固定不变时，**Binom_Dist** 方法适用于固定次数的检验或试验。例如，**Binom_Dist** 方法可以计算出生的三个婴儿中两个是男孩的概率。



Number_s 和 trials 将被截尾取整。
如果 number_s、trials 或 probability_s 为非数值，则 Binom_Dist 方法将生成错误。
如果 number_s < 0 或 number_s > trials，则 Binom_Dist 方法将生成错误。
如果 probability_s < 0 或 probability_s > 1，则 Binom_Dist 方法将生成错误。
二项式概率密度函数的计算公式为：

其中：

![img]()

为 COMBIN(n,x)。

累积二项式分布的计算公式为：

 

#### **WorksheetFunction.Binom_Inv**

返回一元二项式概率分布的反函数。

**语法**

**express.Binom_Inv(Arg1, Arg2, Arg3)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                             |
| -------- | ------------- | ------------ | ------------------------------------ |
| *Arg1*   | 必选          | **Double**   | Trials - 贝努利试验次数。            |
| *Arg2*   | 必选          | **Double**   | Probability_s - 每次试验成功的概率。 |
| *Arg3*   | 必选          | **Double**   | Alpha - 临界值。                     |

**返回值**

Double

**说明**

- 如果 Trials、Probability_s 或 Alpha 为非数值型，**Binom_Inv** 方法将生成错误。
- 如果 Trials 不是整数，则截尾取整。
- 如果 Trials < 0，则 **Binom_Inv** 方法生成错误。
- 如果 Probability_s < 0 或 Probability_s > 1，则 **Binom_Inv** 方法生成错误。
- 如果 Alpha < 0 或 Alpha > 1，则 **Binom_Inv** 方法生成错误。

#### **WorksheetFunction.BinomDist**

返回一元二项式分布的概率。

**语法**

**express.BinomDist(Arg1, Arg2, Arg3, Arg4)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Double**   | 试验成功的次数。                                             |
| *Arg2*   | 必选          | **Double**   | 独立试验的次数                                               |
| *Arg3*   | 必选          | **Double**   | 每次试验成功的概率。                                         |
| *Arg4*   | 必选          | **Boolean**  | 一个逻辑值，用于确定函数的形式。如果 cumulative 为 TRUE，则 BinomDist 返回累积分布函数，即最多存在 number_s 次成功的概率；如果为 FALSE，则返回概率密度函数，即存在 number_s 次成功的概率。 |

**返回值**

Double

**说明**

当任何试验的结果仅包含成功或失败两种情况，当试验是独立试验，以及当整个试验过程中成功的概率固定不变时，BinomDist 适用于固定次数的检验或试验。例如，BinomDist 可以计算出生的三个婴儿中两个是男孩的概率。

- Number_s 和 trials 将被截尾取整。
- 如果 number_s、trials 或 probability_s 为非数值型，则 BinomDist 将生成一个错误。
- 如果 number_s < 0 或 number_s > trials，则 BinomDist 将生成一个错误。
- 如果 probability_s < 0 或 probability_s > 1，则 BinomDist 将生成一个错误。
- 二项式概率密度函数的计算公式为：

其中：

![img]()

为 COMBIN(n,x)。

累积二项式分布的计算公式为：



#### **WorksheetFunction.Ceiling**

返回向上舍入（远离零）到最接近的 significance 的倍数的 number。

**要点**??此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 **Ceiling_Precise** 方法。

**语法**

**express.Ceiling(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                |
| -------- | ------------- | ------------ | --------------------------------------- |
| *Arg1*   | 必选          | **Double**   | Number - 要舍入的值。                   |
| *Arg2*   | 必选          | **Double**   | Significance - 用以进行舍入计算的倍数。 |

**返回值**

Double

**说明**

例如，如果要避免在价格中使用“分”，而产品价格为 ￥4.42, 那么可使用公式 `Ceiling(4.42,0.05)` 将价格进位到最近的“角”。

- 如果任一参数为非数值，**Ceiling** 将生成错误。
- 不论 number 的符号如何，向远离零的方向调整时，值都会向上舍入。如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。
- 如果 number 和 significance 的符号不同，则 **Ceiling** 将生成错误。

#### **WorksheetFunction.Ceiling_Precise**

返回舍入到最接近的有效位倍数的指定数字。

**语法**

**express.Ceiling_Precise(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                |
| -------- | ------------- | ------------ | --------------------------------------- |
| *Arg1*   | 必选          | **Double**   | Number - 要舍入的值。                   |
| *Arg2*   | 可选          | **Variant**  | Significance - 用以进行舍入计算的倍数。 |

**返回值**

Double

**说明**

例如，如果要避免在价格中使用“分”，而产品价格为 ￥4.42, 那么可使用公式 `Ceiling(4.42,0.05)` 将价格进位到最近的“角”。

根据 number 和 significance 参数的符号，**Ceiling_Precise** 方法离零舍入或向零舍入。

| 符号 (Arg1/Arg2) | 舍入       |
| ---------------- | ---------- |
| -/-              | 向零舍入。 |
| +/+              | 离零舍入。 |
| -/+              | 向零舍入。 |
| +/-              | 离零舍入。 |

- 如果任一参数为非数值，**Ceiling_Precise** 将生成错误。
- 如果 number 正好是 significance 的倍数，则无需进行任何舍入处理。

#### **WorksheetFunction.ChiDist**

返回 χ2 分布的单尾概率。

此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 ChiSq_Dist_RT 和 ChiSq_Dist 方法。

**语法**

**express.ChiDist(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Arg1*   | 必选          | **Double**   | 用来计算分布的值。 |
| *Arg2*   | 必选          | **Double**   | 自由度数。         |

**返回值**

Double

**说明**

χ2 分布与 χ2 检验相关。使用 χ2 检验可以比较观察值和期望值。

例如，某项遗传学实验可能假设下一代植物将呈现出某一组颜色。通过对观察到的结果和所期望的结果进行比较，可以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiDist 将生成一个错误。
- 如果 x 为负数，则 ChiDist 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1 或 degrees_freedom > 10^10，则 ChiDist 将生成一个错误。
- ChiDist 按 ChiDist =P(X>x) 计算，其中 X 为 χ2 随机变量。



#### **WorksheetFunction.ChiInv**

返回 χ2 分布的单尾概率的反函数。

此函数已被一个或多个新函数取代，这些新函数可以提供更高的准确度，而且它们的名称可以更好地反映出其用途。仍然提供此函数是为了保持与 ET 早期版本的兼容性。但是，如果不需要后向兼容性，则应考虑从现在开始使用新函数，因为它们可以更加准确地描述其功能。

有关新函数的详细信息，请参阅 ChiSq_Inv_RT 和 ChiSq_Inv 方法。

**语法**

**express.ChiInv(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Arg1*   | 必选          | **Double**   | 与 χ2 分布相关的概率。 |
| *Arg2*   | 必选          | **Double**   | 自由度数。             |

**返回值**

Double

**说明**

如果 probability = ChiDist(x,...)，则 ChiInv(probability,...) = x。使用此函数可对观测到的结果和所期望的结果进行比较，以判定初始假设是否有效。

- 如果两个参数中的任意一个为非数值型，则 ChiInv 将生成一个错误。
- 如果 probability < 0 或 probability > 1，则 ChiInv 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1 或 degrees_freedom ≥ 10^10，则 ChiInv 将生成一个错误。

如果已给定概率值，则 ChiInv 将使用 ChiDist(x, degrees_freedom) = probability 求解值 x。因此，ChiInv 的精度取决于 ChiDist 的精度。ChiInv 使用迭代搜索技术。如果搜索在 64 次迭代之后没有收敛，则该函数将生成一个错误。 

#### **WorksheetFunction.ChiSq_Dist**

返回 χ2 分布。

**语法**

**express.ChiSq_Dist(Arg1, Arg2, Arg3)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 必选          | **Double**   | x - 要在计算分布时使用的值。                                 |
| *Arg2*   | 必选          | **Double**   | deg_freedom - 自由度数。                                     |
| *Arg3*   | 可选          | **Variant**  | cumulative - 一个决定函数形式的逻辑值。如果 cumulative 为 TRUE，则 CHISQ_DIST 将返回累积分布函数；如果为 FALSE，则将返回概率密度函数。 |

**返回值**

Double

**说明**

- 如果任一参数为非数值型，则 CHISQ_DIST 将返回错误值 #VALUE!。
- 如果 x 为负数，则 CHISQ_DIST 将返回错误值 #NUM!。
- 如果 deg_freedom 不是整数，则将被截尾取整。



#### **WorksheetFunction.ChiSq_Dist_RT**

返回 χ2 分布的右尾概率。

**语法**

**express.ChiSq_Dist_RT(Arg1, Arg2)**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Arg1*   | 必选          | **Double**   | 用来计算分布的值。 |
| *Arg2*   | 必选          | **Double**   | 自由度数。         |

**说明**

χ2 分布与 χ2 检验相关。使用 χ2 检验可以比较观察值和期望值。

例如，某项遗传学实验可能假设下一代植物将呈现出某一组颜色。通过对观察到的结果和所期望的结果进行比较，可以判定最初的假设是否有效：

- 如果任一参数为非数值型，则 ChiSq_Dist_RT 将生成一个错误。
- 如果 x 为负数，则 ChiSq_Dist_RT 将生成一个错误。
- 如果 degrees_freedom 不是整数，则将被截尾取整。
- 如果 degrees_freedom < 1 或 degrees_freedom > 10^10，则 ChiSq_Dist_RT 将生成一个错误。
- ChiSq_Dist_RT 的计算公式为 ChiSq_Dist_RT = P(X>x)，其中 X 为 χ2 随机变量。



**成员属性**

#### **WorksheetFunction.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象。可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序。只读。

**语法**

**express.Application**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") } }` |

#### **WorksheetFunction.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **WorksheetFunction** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **WorksheetFunction.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **WorksheetFunction** 对象的变量。

适用环境：web

适用平台：windows/linux