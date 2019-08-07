# NumPadUltra
笔记本小键盘输入扩展（键盘钩子）

## 功能
目前主流笔记本个人电脑为节省键盘占用空间，一般不搭载数字小键盘；但是在一些情景下，小键盘按键的功能并不能被上方横排数字键所完全替代。例如：某些设计软件（如Adobe系列）和大型单机游戏的快捷键必须用到小键盘按键；以及在统计时、需要输入大量数据的情况，小键盘的九宫格布局可大幅加快输入速度。

本程序的开发动机主要是P大的一门名为“物理化学实验”的必修课：

* 经常要向Excel/Origin里输入几页的原始数据，用上方横排数字键实在太慢；
* 要在实验记录本里抄电子版讲义的原理/步骤等内容，本子压在键盘上写字容易误触按键。

本程序的主要功能便是使用（全局）键盘钩子实现了：

* 将笔记本电脑键盘的右半部分临时开辟出来作为数字小键盘的替代按键（称为`小键盘模式`），例如按下`<,`，`>.`，`?/`键后，实际输入的是`Num 1`，`Num 2`，`Num 3`（详见下图）；
* 锁定键盘（即令所有键盘事件无效化），防止误触（称为`锁定模式`）；
* `正常模式`（保留原始键盘布局和功能）和`小键盘模式`、`锁定模式`间可快速切换。

<p align="center">
  <img src="/Keybd.jpg" width="70%" height="70%"/>
  <br>小键盘模式下的实际键盘布局
</p>

## 使用方法
下载最新版本的[单个可执行文件](https://github.com/Z-H-Sun/NumPadUltra/releases/download/v3.26/NumPadUltra.exe)至任意位置，双击运行即可。

* 程序刚启动时为`小键盘模式`。默认在屏幕的左上角会出现下图所示的半透明图标，单击拖拽可改变其位置；<p align="center"><img src="/NumPadUltra/KeyboardHook/KeybdNum.ico" width="32" height="32"/></p>现在可以尝试开一个记事本，依次按下键盘上的`op[l;',./`，如果程序正常运行，实际显示的将是`123456789`
* 按下右Windows徽标键，或右键单击上述图标，可在`锁定模式`和`小键盘模式`之间切换；图标的颜色分别变为红色、黄色；<p align="center"><img src="/NumPadUltra/KeyboardHook/KeybdFbd.ico" width="32" height="32"/></p>在`锁定模式`下，计算机不会响应除了Windows徽标键以外的任何键盘事件
* 按下左Windows徽标键，或左键单击上述图标，可在上述的某个模式与`正常模式`之间切换；图标的颜色分别变为红或黄色/绿色；<p align="center"><img src="/NumPadUltra/KeyboardHook/Keybd.ico" width="32" height="32"/></p>
* 若长按Windows徽标键（长于1秒），则实现按下Windows键本身的功能，不切换模式；
* 双击上述图标可退出程序
* 鼠标中键单击上述图标，可显示如下图所示的键盘布局示意图：<p align="center"><img src="/NumPadUltra/KeyboardHook/Keyboards%20Collection.png" width="70%" height="70%"/></p>
    * 单击拖拽可改变位置，滑动滚轮以改变透明度
    * 点击ESC按钮或按下ESC以关闭窗口
    * 在相应的按键上悬停以查看实际键盘布局，单击可模拟`小键盘模式`中相应的键盘按键事件

* Q: 是否可以自定义“数字键盘布局”？A: 目前暂时无法进行配置，只能手动更改[Form1.vb](/NumPadUltra/KeyboardHook/Form1.vb)中的`KeyBoardProc`函数，然后自行编译