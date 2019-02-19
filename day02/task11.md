### 单元格的选取

1.表示一个单元格

![示一个单元](/Users/zyz/Desktop/myDesktop/VBA/day02/表示一个单元格.png)

2.表示相邻单元格区域

![示相邻单元格区](/Users/zyz/Desktop/myDesktop/VBA/day02/表示相邻单元格区域.png)

offset第一个参数是上下(上正下负)偏移量，第二个参数是左右(左正右负)偏移量

range("a1:a10").offset(0,1):表示range("a1:b10") 这一区间

resize(偏移的总行数，偏移的总例数)

3.表示不相邻的单元格区域

![示不相邻的单元格区域](/Users/zyz/Desktop/myDesktop/VBA/day02/表示不相邻的单元格区域1.png)

union demo选取a列的偶数行

![示不相邻的单元格区域](/Users/zyz/Desktop/myDesktop/VBA/day02/表示不相邻的单元格区域2.png)

![幕快照 2019-02-06 上午11.19.5](/Users/zyz/Desktop/myDesktop/VBA/day02/屏幕快照 2019-02-06 上午11.19.58.png)

4.表示行

![示](/Users/zyz/Desktop/myDesktop/VBA/day02/表示行.png)

选取连续的行用Rows，选取不连续的行用Range

EntireRow表示区域所在的行数

5.表示列

![示](/Users/zyz/Desktop/myDesktop/VBA/day02/表示列.png)



EntireColumn表示区域所在的列数

6.重置坐标下的单元格表示方法

![置坐标下的单元格表示方](/Users/zyz/Desktop/myDesktop/VBA/day02/重置坐标下的单元格表示方法.png)

本来是以a1为顶点的改为以b2为顶点，再在b2里面赋值100

7.表示正在选取的单元格区域

![示正在选取的单元格区](/Users/zyz/Desktop/myDesktop/VBA/day02/表示正在选取的单元格区域.png)

给所选中的区域都填充100

练习作业

![练习作业](/Users/zyz/Desktop/myDesktop/VBA/day02/练习作业.png)