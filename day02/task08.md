### vba分支与END语句

#### END

一 END语句

![nd语](/Users/zyz/Desktop/myDesktop/VBA/day02/end语句.png)

END和exit sub的区别

![体end和exit sub的代](/Users/zyz/Desktop/myDesktop/VBA/day02/窗体end和exit sub的代码.png)

![体end和exit sub的区](/Users/zyz/Desktop/myDesktop/VBA/day02/窗体end和exit sub的区别.png)

end会将窗体也退出掉，而exit sub只会退出当前程序

二Exit语句 强制退出

![xit语](/Users/zyz/Desktop/myDesktop/VBA/day02/exit语句.png)

1.Exit sub 退出当前这个程序

![xit su](/Users/zyz/Desktop/myDesktop/VBA/day02/exit sub.png)

2.Exit function 退出当前这个函数

![xit functio](/Users/zyz/Desktop/myDesktop/VBA/day02/exit function.png)

3.Exit for 退出当前的for循环

![xit fo](/Users/zyz/Desktop/myDesktop/VBA/day02/exit for.png)

4.Exit do 退出当前的do循环

![xit d](/Users/zyz/Desktop/myDesktop/VBA/day02/exit do.png)

#### 分支

1.goto 

demo：强制用户输入

![oto语](/Users/zyz/Desktop/myDesktop/VBA/day02/goto语句.png)

inputbox图标

![oto dem](/Users/zyz/Desktop/myDesktop/VBA/day02/goto demo.png)

点击取消返回false，len(sr)=5

2.gosub…return跳过去了再回来 return返回到gosub后面

![oSub retur](/Users/zyz/Desktop/myDesktop/VBA/day02/goSub return.png)

demo：a列里面是偶数的将原来单元格的值改为偶数

![幕快照 2019-02-05 下午8.00.4](/Users/zyz/Desktop/myDesktop/VBA/day02/屏幕快照 2019-02-05 下午8.00.46.png)

![幕快照 2019-02-05 下午8.01.0](/Users/zyz/Desktop/myDesktop/VBA/day02/屏幕快照 2019-02-05 下午8.01.09.png)

3.on error resume next 遇到错误，跳过执行下一句

![n error rusume nex](/Users/zyz/Desktop/myDesktop/VBA/day02/on error rusume next.png)

demo:将A列和B列的数字相乘

![emo运行](/Users/zyz/Desktop/myDesktop/VBA/day02/demo运行前.png)

字母和数字类型不一样，不能相乘，会报错

![错窗](/Users/zyz/Desktop/myDesktop/VBA/day02/报错窗口.png)

取消掉on error resume next后的执行结果

![行后的效](/Users/zyz/Desktop/myDesktop/VBA/day02/执行后的效果.png)

4.on error goto出错时调到指定行数

![n error got](/Users/zyz/Desktop/myDesktop/VBA/day02/on error goto.png)

MsgBox

![sgBo](/Users/zyz/Desktop/myDesktop/VBA/day02/msgBox.png)

5.on error goto 0 取消错误跳转

![n error goto ](/Users/zyz/Desktop/myDesktop/VBA/day02/on error goto 0.png)

![行的程序结](/Users/zyz/Desktop/myDesktop/VBA/day02/执行的程序结果.png)

程序在x=8的时候出错

![幕快照 2019-02-05 下午8.17.0](/Users/zyz/Desktop/myDesktop/VBA/day02/屏幕快照 2019-02-05 下午8.17.04.png)





















