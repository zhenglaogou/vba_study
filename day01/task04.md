## 循环语句

Demo2:计算总金额

![demo2](/Users/zyz/Desktop/myDesktop/VBA/day01/demo2.png)

for语句

![for循环](/Users/zyz/Desktop/myDesktop/VBA/day01/for循环.png)

~~~
for x = 2 to 6 step 1
for x = 1000 to 10 step -10
~~~

each in语句

![each in循环](/Users/zyz/Desktop/myDesktop/VBA/day01/each in循环.png)

Offset(0, -1):纵坐标不变，横坐标向左移动1个单位

Offset(1, 0):横坐标不变，纵坐标向下移动1个单位

do loop语句

![do loop循环](/Users/zyz/Desktop/myDesktop/VBA/day01/do loop循环.png)



do loop变种

![do loop变种](/Users/zyz/Desktop/myDesktop/VBA/day01/do loop变种.png)



Demo3(for each in语句的典范):将两个区域中为空的单元格填充为0

![屏幕快照 2019-02-05 上午10.47.52](/Users/zyz/Desktop/myDesktop/VBA/day01/屏幕快照 2019-02-05 上午10.47.52.png)

![demo3 each in ](/Users/zyz/Desktop/myDesktop/VBA/day01/demo3 each in .png)

当循环有规律的时候用for循环，当循环没有规律的时候用for each in循环

do loop demo：判断是否是连续的数字，如果不是在右边写上断点。

![do loop demo](/Users/zyz/Desktop/myDesktop/VBA/day01/do loop demo.png)