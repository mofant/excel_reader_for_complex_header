## 总体目标
---
+ 提取execl的大标题，如果excel存在的话。
+ 把复杂表头转成一维表头。
+ 把左侧的合并的索引展开变成多维。

### 表头和数据拆分
结合之前的经验，先对表格做表头和数据的拆分。
表头和表数据的拆分依据分为：
+ 自某一行开始，数据块所在的列数据格式统一。
+ 自某一行开始，数据块中没有合并单元格。
+ 或自某一行开始，其所有的单元格都是有内外边框。

数据格式统一：
+ 非空单元格，其格式统一，如数值和可以转换成数值的文本是同一类型。


表头与数据划分
+ 搜索没有合并单元格的区域，（假定为右下方）。
+ 搜索有内外边框的单元格如果都没有单元格，则忽略。
+ 搜索每一列的没有合并单元格的单元格数据格式，找出数据列相同行号。
>> 结合上面三点获取数据列和表头列。


## 表头拼接
+ 假定每一列都有列名，存在于合并和未合并的单元格中。

表头合并单元格值处理方式
+ 在横向合并单元格中，其每一列的都将有第一个单元格的值。
+ 在竖向合并中，只保留开始列的值。
