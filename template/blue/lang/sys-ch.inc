﻿<%
Dim word(800)
rem 已使用到653



word(0)="目前暂无任何信息" 
word(1)="招聘进行中"
word(2)="招聘已结束"
word(3)="隐蔽内容,会员登陆可查看"
word(4)="此留言暂无回复！"
word(495)="暂无回复"
rem 搜索框
word(5)="关键词不能为空"
word(6)="请选择分类"
word(7)="请输入关键词"
word(8)="请输入新闻关键词"
word(9)="请输入产品关键词"
word(10)="分类"
word(11)="新闻"
word(12)="产品"
word(13)="图片"
word(14)="视频"
word(15)="下载"
rem 搜索结果
word(16)="搜索到符合"
word(17)="的结果约"
word(18)="个"
word(19)="抱歉！搜索不到任何关于"
word(20)="的信息"

rem 邮件订阅
word(21)="填写Email订阅本站最新动态"
word(22)="请检查Email是否正确"

rem 订单查询页面
word(23)="请输入您的订单号"
word(24)="还未回复"
word(30)="等待货款"
word(31)="已支付"
word(32)="处理中"
word(33)="已发货"
word(34)="订单完成"
word(41)="无关联此订单号的订单!请查看订单号是否正确!"

rem 上条信息，下条信息 没有时
word(42)="无"

rem 自定义表单
word(43)="此表单没有关联的ID"
word(44)="此表单已经关闭，请联系管理员！"
word(45)="不能为空"
word(46)="邮箱格式不正确"
word(47)="身份证号请填写15位或者18位"
word(48)="必须为数字"
word(49)="手机号格式不正确"
word(50)="邮政编码格式不正确"
word(51)="验证码："
word(52)="验证码不能为空"
word(493)="验证码"
word(494)="请输入验证码"
word(53)="提交"
word(54)="重置"
word(55)="关闭"
rem 投票结果页面

word(58)="没有此投票"
word(59)="请选择投票项目"
word(60)="该投票还没开始"
word(61)="该投票已经过期"
word(62)="恭喜您投票成功"
word(63)="请勿重复投票"
word(64)="没有此投票"



rem 留言私密功能
word(69)="先生"
word(70)="会员功能"
word(71)="#ff6600"
rem  发布留言
word(72)="姓名："
word(73)="姓名不能为空"
word(74)="请输入您的真实姓名"
word(75)="联系电话："
word(76)="联系电话不能为空"
word(77)="手机号格式不正确"
word(78)="请输入您常用的联系电话"
word(79)="电子邮件："
word(80)="电子邮件不能为空"
word(81)="电子邮件格式不正确"
word(82)="请输入您常用的电子邮件"
word(83)="联系QQ："
word(84)="联系QQ不能为空"
word(85)="必须为数字"
word(86)="请输入您常用的QQ号码"
word(87)="咨询内容："
word(88)="咨询内容不能为空"
word(89)="请输入您要咨询的内容"
word(90)="悄悄话："
word(91)="验证码："
word(92)="验证码不能为空"
word(93)="提交"
word(94)="重置"
rem 会员中心个人资料
word(95)="请登陆您的会员账号"
word(96)="基本资料"
word(97)="登录名"
word(98)="所属组别"
word(99)="真实姓名"
word(100)="性别"
word(512)="先生"
word(513)="女士"
word(514)="空表示不修改"
word(101)="设置密码"
word(102)="确认密码"
word(103)="密码提示问题"
word(104)="密码提示答案"
word(105)="单位名称"
word(106)="地址"
word(107)="邮编"
word(108)="传真"
word(109)="电话"
word(110)="电子邮箱"
word(111)="验证码"
word(112)="注册时间"
word(113)="登陆次数"
word(114)="最后登陆时间"
word(115)="最后登陆IP"
word(116)="保存"
word(117)="重置"
rem 会员注册
word(118)="会员禁止注册，请联系管理员！"
word(119)="用户名："
word(120)="用户名不能为空"
word(121)="用户名长度为6-16个任意字符组成"
word(122)="由0-9,a-z,-_任意组合6-16个的字符串"
word(123)="设置密码："
word(124)="密码至少6个字符,最多18个字符！"
word(125)="密码不能为空"
word(126)="密码强度："
word(127)="弱"
word(128)="一般"
word(129)="强"
word(130)="密码由6-18位任意字符组成！"
word(131)="确定密码："
word(132)="两次输入的密码不一致！"
word(133)="确认密码不能为空"
word(134)="请再一次输入密码"
word(135)="提示问题："
word(136)="密码提示问题不能为空"
word(137)="请选择您的密码提示问题"
word(138)="我妈妈的名字是？"
word(139)="我就读的初中叫什么名字？"
word(140)="我就读的小学叫什么名字？"
word(141)="我就读的幼儿园是？"
word(142)="我的故乡是？"
word(143)="我最喜欢的颜色？"
word(144)="我最喜欢的数字？"
word(145)="用于找回密码时使用"
word(146)="提示答案："
word(147)="密码提示答案不能为空"
word(148)="请牢记您的提示答案"
word(149)="用户姓名："
word(150)="用户姓名不能为空"
word(151)="姓名长度为2-6位字符"
word(152)="请填写您的真实姓名"
word(153)="用户性别："
word(154)="先生"
word(155)="女士"
word(156)="单位名称："
word(157)="用户地址："
word(158)="邮政编码："
word(159)="传真号码："
word(160)="手机号码："
word(161)="手机号格式不正确"
word(162)="常用手机号码不能为空"
word(163)="请填写您常用的手机号码"
word(164)="电子邮箱："
word(165)="电子邮箱格式不正确"
word(166)="常用电子邮箱不能为空"
word(167)="请填写您常用的电子邮箱"
word(168)="验证码："
word(169)="验证码不能为空"
word(170)="注册"
word(171)="重置"

rem  会员找回密码
word(172)="用户名："
word(173)="用户名不能为空"
word(174)="用户名长度为6-16个任意字符组成"
word(175)="请输入您注册时的用户名"
word(176)="电子邮箱："
word(177)="电子邮箱格式不正确"
word(178)="常用电子邮箱不能为空"
word(179)="请输入您注册时的电子邮箱"
word(180)="下一步"
word(181)="登录名和电子邮箱不对应，"
word(182)="登录名："
word(183)="电子邮箱："
word(184)="密码提示问题："
word(185)="密码提示答案："
word(186)="答案不能为空"
word(187)="请输入您正确的答案"
word(188)="您在"
word(189)="的会员密码已经找回！"
word(190)="以下是您的密码信息！请登陆以后修改成安全,易记的密码！"
word(191)="用户名："
word(192)="密码："
word(193)="本邮件由系统自动发送,无须回复"
word(194)="网址："
word(195)="邮箱："
word(196)="电话："
word(197)="传真："
word(198)="地址："
word(199)="新密码已发送到您的邮箱里了，请登陆您的邮箱查看！"
word(515)="您的新密码是："
word(516)="请登录以后修改您的密码！"



rem  会员中心  我的留言
word(200)="暂无留言"
word(201)="留言内容"
word(202)="隐秘内容"
word(203)="管理员回复"
word(204)="此留言暂未回复"
rem  会员中心  我的应聘列表
word(205)="应聘职位"
word(206)="姓名"
word(207)="联系电话"
word(208)="应聘日期"
word(209)="状态"
word(210)="操作"
word(211)="已回复"
word(212)="未回复"
word(213)="查看详情"
word(214)="暂无应聘信息"

rem  会员中心  我的订单列表
word(234)="还未回复"
word(236)="支付宝支付"
word(237)="银行转账"

word(239)="等待货款"
word(240)="已支付"
word(241)="处理中"
word(242)="已发货"
word(243)="订单完成"
word(249)="支付"
word(250)="确认收货"
word(251)="已确认"
word(253)="暂无订单信息"

rem  会员中心  我的订单详情

word(269)="支付宝支付"
word(270)="银行汇款"
word(278)="等待货款"
word(279)="已支付"
word(280)="处理中"
word(281)="已发货"
word(282)="订单完成"


rem 当前位置
word(297)="所有信息"
word(298)="网站公告"
word(299)="推荐信息"
word(300)="首页"

rem  提交简历
word(301)="非常抱歉！非本站会员不可以在线应聘！"
word(302)="应聘职位："
word(303)="应聘职位保持默认即可"
word(304)="姓名："
word(305)="姓名不能为空"
word(306)="姓名长度为2-6位字符"
word(307)="请输入您的真实姓名"
word(308)="性别："
word(309)="先生"
word(310)="女士"
word(311)="婚姻状况："
word(312)="未婚"
word(313)="已婚"
word(314)="出生日期："
word(315)="出生日期不能为空"
word(316)="格式如：1986-12-16"
word(317)="身高："
word(318)="身高不能为空"
word(319)="必须为数字"
word(320)="格式如：178 单位为cm"
word(321)="毕业院校："
word(322)="毕业院校不能为空"
word(323)="请输入您的毕业院校"

word(324)="学历："
word(325)="学历不能为空"
word(326)="请输入您的学历"
word(327)="专业："
word(328)="专业不能为空"
word(329)="请输入您的专业"
word(330)="毕业时间："
word(331)="毕业时间不能为空"
word(332)="请输入您的毕业时间"
word(333)="户籍地："
word(334)="户籍地不能为空"
word(335)="请输入您的户籍地"
word(336)="教育经历："
word(337)="教育经历不能为空"
word(338)="请输入您的教育经历"
word(339)="工作经历："
word(340)="工作经历不能为空"
word(341)="请输入您的工作经历"
word(342)="手机号码："
word(343)="手机号码格式不正确"
word(344)="手机号码不能为空"
word(345)="请输入您常用的手机号"
word(346)="邮箱地址："
word(347)="电子邮箱格式不正确"
word(348)="电子邮箱不能为空"
word(349)="请输入您的电子邮箱"
word(350)="通信地址："
word(351)="通信地址不能为空"
word(352)="请输入您的通信地址"
word(353)="邮政编码："
word(354)="邮编号码不能为空"
word(355)="邮政编码格式不正确"
word(356)="请输入您的邮政编码"
word(357)="验证码："
word(358)="验证码不能为空："
word(359)="提交"
word(360)="重置"

rem ajax 验证
word(361)="留言内容请不要带有攻击人的语言"
word(362)="通过信息验证"
word(363)="您的用户不存在"
word(364)="您的电子邮箱不存在"
word(365)="密码提示问题不存在"
word(366)="已存在请更换用户名"
word(367)="已存在请更换电子邮箱"
word(368)="验证码不正确"
word(648)="正在检测信息..."


rem 商品规格无选择时提示
word(369)="请选择您要购买的商品规格！"

word(370)="不能在5秒内重新打开此页!"

rem 自定义表单提交信息

word(371)="信息提交成功"

rem  修改会员资料验证页面
word(372)="设置密码长度应为6-18个任意字符串"
word(373)="设置密码和确定密码不一致"
word(374)="单位名称、地址长度不能超过100个字符"
word(375)="邮编长度不能超过20个字符"
word(376)="常用电话不能超过11个字符"
word(377)="电子邮箱格式不正确"
word(378)="电子邮箱已经被注册过"
word(379)="验证码错误或已失效"
word(380)="修改会员资料成功"

rem 会员注册验证页面

word(381)="您在"
word(382)="注册成功！"
word(383)="恭喜您注册成为"
word(384)="的会员，谢谢您的支持！以下是您的注册信息！请牢记！"
word(385)="用户名："
word(386)="密码："
word(387)="密码保护问题："
word(388)="密码保护问题答案："
word(389)="本邮件由系统自动发送,无须回复"
word(390)="网址: "
word(391)="邮箱: "
word(392)="电话："
word(393)="传真："
word(394)="地址： "
word(395)="注册成功,请等待管理员的审核!,注册信息已经发往您的邮箱！"
word(396)="注册成功,请等待管理员的审核"
word(397)="注册成功！注册信息已经发往您的邮箱！"
word(398)="注册成功！"
word(596)="绑定成功,请等待管理员的审核!,注册信息已经发往您的邮箱！"
word(597)="绑定成功,请等待管理员的审核"
word(598)="绑定成功！注册信息已经发往您的邮箱！"
word(600)="绑定成功！"

rem 手机版留言注册验证

word(399)="留言标题不能为空！"
word(400)="标题中请不要带有攻击人的语言"
word(401)="留言内容中请不要带有请不要带有攻击人的语言"
word(402)="qq不能为空！"
word(403)="电话号码必填具不能少于11个字符！"
word(404)="电子邮箱必填具不能少于6个字符！"
word(405)="留言内容不能为空"
word(406)="一条来自"
word(407)="的新留言"
word(408)="以下是关于"
word(409)="留言的详细信息！"
word(410)="联系电话："
word(411)="电子邮箱："
word(412)="联系QQ:"
word(413)="留言内容:"
word(414)="留言提交成功,经过我们审核后在前台可以显示"
word(415)="您的留言提交成功，系统会将您的留言信息发送邮件提醒管理员！以尽快查看您的留言信息！处理的信息将会发送到您的邮箱！请随时关注您的邮箱"
word(416)="留言提交成功,工作人员正在处理您的留言!"
word(443)="留言提交成功!正在审核中..."
word(444)="留言提交成功!已发送邮件通知管理员！"
word(445)="留言提交成功!工作人员会尽快回复您的留言！"
rem 手机版提交简历验证页面

word(417)="一条来自"
word(418)="应聘"
word(419)="信息"
word(420)="以下是关于"
word(421)="的应聘详细信息！"
word(422)="应聘职位："
word(423)="姓名："
word(424)="性别："
word(425)="婚姻状况："
word(426)="出生日期："
word(427)="身高："
word(428)="毕业院校："
word(429)="学历："
word(430)="专业:"
word(431)="毕业时间:"
word(432)="户籍地："
word(433)="教育经历："
word(434)="工作经历："
word(435)="联系电话："
word(436)="电子邮箱："
word(437)="通信地址："
word(438)="邮政编码："
word(439)="本邮件由"
word(440)="自动发送,无须回复"
word(441)="您的简历提交成功，系统会将您的应聘信息发送邮件提醒管理员！以尽快查看您的应聘信息！处理的信息将会发送到您的邮箱！请随时关注您的邮箱！"
word(442)="您的简历提交成功，如果合适我们会尽快通知您面试!"

word(459)="简历提交成功!已发送邮件通知管理员！"
word(460)="简历提交成功!工作人员会尽快回复您"

rem 购物车验证页面


word(446)="您的购物车里没有商品"
word(447)="来自"
word(448)="的"
word(449)="信息"
word(450)="订单号："
word(451)="性别："
word(452)="单位名称："
word(453)="地址："
word(454)="邮编："
word(455)="传真："
word(456)="邮箱地址："
word(457)="联系电话："
word(458)="订单提交成功"


rem  邮件订阅验证
word(461)="此信箱已经订阅！请更换其他邮箱！"
word(462)="您的信箱"
word(463)="订阅信息成功！"
word(464)="退订成功！欢迎您再次订阅！"
word(465)="此信箱还没有订阅！欢迎您订阅！"

rem 在线客服

word(466)="QQ在线客服"
word(467)="旺旺在线客服"

rem 会员登录验证
word(468)="退出成功！正在跳转中......"
word(469)="登录名不正确，请返回。"
word(470)="登录密码不正确，请返回。"
word(471)="帐号非使用状态，请返回。"
word(472)="帐号正在审核中...，请返回。"
word(473)="验证码错误或不能为空，请返回。"
word(474)="登陆成功！正在跳转中......"



rem 列表页分页PC端 
rem 475 表示分页样式 默认值：default 目前支持molv、yahei、flow 也支持16进制颜色值 例如：#ff0000
rem 476  用于控制首页。值支持三种类型。如：first: '首页' 如：first: false，则表示不显示首页项
rem 477 478 表示 用于控制上一页 和下一页。若不显示，设置false即可
rem 479 表示 用于控制尾页。值支持三种类型如：last: '尾页' 如：last: false，则表示不显示首页项
rem 480  表示连续分页数
rem 481 是否显示跳转
word(475)="#1970b8"
word(476)="1"
word(477)="'上一页'"
word(478)="'下一页'"
word(479)=""
word(480)="5"
word(481)="false"

rem 列表页手机版分页
word(482)="#1970b8"
word(483)="1"
word(484)="'<'"
word(485)="'>'"
word(486)=""
word(496)="1"
word(497)="false"


rem 内容页分页
word(498)="#1970b8"
word(499)="false"
word(500)="'上一页'"
word(501)="'下一页'"
word(502)="false"
word(503)="5"
word(504)="false"

rem 内容页手机版分页
word(505)="#1970b8"
word(506)="false"
word(507)="'<'"
word(508)="'>'"
word(509)="false"
word(510)="1"
word(511)="false"


rem 其他
word(517)="全部"

rem  内容设置权限时的提示语  infoquanxian.asp
word(531)="您所在的用户组无法查看"


word(622)="如果您的浏览器没有自动跳转，请点击这里"
word(623)="您的购买价格："
rem 手机 留言和提交简历处理页面
word(624)="确定"

rem 加入购物车提示

word(518)="追梦工作室提示您"
word(519)="去结算"
word(520)="继续购物"
word(521)="已成功添加到购物车内!"
word(522)="商品总计："
word(523)="件"
word(524)="总计："

rem showcart.asp

word(525)="您已选择："
word(526)="库存数量为："
word(527)="已销售："
word(528)="规格为"
word(529)="商品已经售罄！请选择其他商品！"
word(530)="抱歉！无相关的规格商品！请选择其他的商品规格！"


rem 我的购物车
word(532)="我的购物车"
word(533)="1.查看购物车"
word(534)="2.填写购物信息"
word(535)="3.完成订单"

word(536)="缩略图"
word(537)="产品名称"
word(538)="市场价"
word(539)="购买价"
word(540)="购买数量"
word(541)="总价"
word(542)="操作"

word(543)="规格："
word(544)="库存："


word(545)="移除"
word(546)="商品总金额(不包含运费)："

word(547)="继续购物"
word(548)="清空商品"
word(549)="提交订单"

word(550)="亲，您的购物车还是空的，您可以："
word(551)="● 去商品挑选喜欢的商品"

word(552)="非本站会员不能购买本站商品！"


word(553)="商品共计："
word(554)="件"
word(555)="商品总金额(不包含运费)："
word(556)="返回上一层"


word(557)="订单说明："
word(558)="产品订购"
word(559)="请输入此订单的说明"

word(560)="订单备注："
word(561)="可对订单进行补充说明"
word(562)="收货人姓名："
word(563)="收货人姓名不能为空"
word(564)="姓名长度为2-6位字符"
word(565)="请输入您的真实姓名"
word(566)="选择性别："
word(567)="先生"
word(568)="女士"
word(569)="单位名称："
word(570)="单位姓名不能为空"
word(571)="个人用户请填写个人购买"
word(572)="收货地址："
word(573)="收货地址不能为空"
word(574)="请输入正确的收货地址！"
word(575)="收货邮编："
word(576)="邮编号码不能为空"
word(577)="邮政编码格式不正确"
word(578)="请填写正确的邮编号码"
word(579)="传真号码："
word(580)="手机号码："
word(581)="手机号格式不正确"
word(582)="手机号码不能为空"
word(583)="请输入您常用的手机号码"
word(584)="电子邮箱："
word(585)="电子邮箱格式不正确"
word(586)="电子邮箱不能为空"
word(587)="请输入您常用的电子邮件"
word(588)="配送方式："
word(589)="配送方式不能为空"
word(590)="请选择您的配送方式"
word(591)="支付方式："
word(592)="请选择您的支付方式"
word(593)="支付方式不能为空"
word(594)="支付宝支付"
word(595)="银行转账"
word(596)="请输入您的支付方式"
word(597)="验证码："
word(598)="验证码不能为空"
word(599)="确认提交"
word(600)="重置订单"







rem 支付页面

word(601)="订单详细"
word(602)="订单编号"
word(603)="[查看详细]"
word(604)="配送方式"
word(605)="商品价格"
word(606)="配送价格"
word(607)="订单总金额"
word(608)="支付信息"
word(609)="支付方式："
word(610)="支付宝支付"
word(611)="银行转账"
word(612)="[更改支付方式]"

word(613)="订单状态："
word(614)="等待货款"
word(615)="已支付"
word(616)="处理中"
word(617)="已发货"
word(618)="订单完成"
word(619)="立即支付"
word(620)="您好！您的身份是游客！如果您不立即支付！请牢记您的订单号和支付页面！"
word(621)="以后支付"

rem  外部提交提示
word(491)="请不要试图直接提交数据,您的IP我们已经记录！"

rem QQ绑定提示
word(492)="解绑成功！请使用您注册的用户名 密码登陆"
word(487)="绑定成功"
word(488)="绑定失败！请检查您的帐户名密码是否正确"
word(489)="绑定成功，账户正在审核中..."
word(490)="账号非使用状态，无法绑定账户"



rem 后台 自动回复留言信件
word(625)="您于"
word(626)="在"
word(627)="的留言已回复"
word(628)="您的留言:"
word(629)="回复内容:"
rem 发送邮件通用文字
word(630)="本邮件由"
word(631)="发送于"
word(632)="网址:"
word(633)="邮箱:"
word(634)="电话:"
word(635)="传真:"
word(636)="地址:"

rem 后台 自动回复应聘信件
word(637)="您于"
word(638)="应聘的"
word(639)="岗位 已答复！"
word(640)="以下是关于您在"
word(641)="应聘的"
word(642)="岗位的详细回复信息"
word(643)="回复于"

rem 后台 自动回复订单信件
word(644)="您于"
word(645)="订购的产品已回复"
word(646)="以下是关于您在"
word(647)="订购产品的的回复信息"

rem sub 页面

word(649)="无标签，请在后台添加标签"
word(650)="无规格"
word(651)="无库存"


rem 支付宝支付成功提示信息

word(652)="支付成功！"
word(653)="请尽快联系我们的客服确认付款信息！"

%>
