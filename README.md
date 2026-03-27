Hi Katy,

Noted with thanks. We will proceed accordingly for this case and reach out if any further checks are needed.

Best regards,
Jason

下面是一份 SHL Code Test 高频题型清单（30题）。
这些题型是根据 SHL 常见题库风格整理的，和 LeetCode / HackerRank 上的题型高度对应。

SHL 的特点：
偏工程逻辑 + 基础算法，很少出现特别复杂的算法。

我按 6个模块整理，每个模块 5题，共 30题。


---

一、Array / HashMap（最常见）

SHL 出现率 最高模块

1 Two Sum

输入数组，找到两个数之和等于 target

核心

hashmap
O(n)


---

2 Top K Frequent Elements

找出现次数最多的 k 个元素

核心

Counter
heap / sort


---

3 Longest Consecutive Sequence

找最长连续数字序列

核心

set
O(n)


---

4 Subarray Sum Equals K

子数组和等于 k 的数量

核心

prefix sum
hashmap


---

5 Group Anagrams

字符串异位词分组

核心

sorted(word)
hashmap


---

二、String 处理（SHL 很喜欢）

6 Valid Parentheses

括号匹配

核心

stack


---

7 Longest Substring Without Repeating

最长无重复字符子串

核心

sliding window


---

8 Longest Palindromic Substring

最长回文子串

核心

expand around center


---

9 Minimum Window Substring

最小覆盖子串

核心

sliding window
hashmap


---

10 String Compression

压缩字符串

例如：

aaabb → a3b2

核心

two pointer


---

三、Sorting / Greedy

11 Merge Intervals

合并区间

核心

sort + scan


---

12 Meeting Rooms

判断会议是否冲突

核心

sort intervals


---

13 Insert Interval

插入新区间并合并

核心

interval logic


---

14 Task Scheduler

任务冷却时间

核心

greedy
heap


---

15 Reorganize String

重排字符串避免相邻相同

核心

max heap


---

四、Sliding Window（SHL高频）

16 Maximum Sliding Window

滑动窗口最大值

核心

deque


---

17 Minimum Size Subarray Sum

最短子数组 >= target

核心

sliding window


---

18 Longest Repeating Character Replacement

核心

window + frequency


---

19 Permutation in String

判断 s2 是否包含 s1 的排列

核心

sliding window
counter


---

20 Fruit Into Baskets

最多两种元素的最长子数组

核心

sliding window


---

五、Graph / BFS（银行很喜欢）

21 Number of Islands

核心

BFS / DFS


---

22 Course Schedule

是否存在课程循环依赖

核心

topological sort


---

23 Shortest Path in Binary Matrix

核心

BFS


---

24 Clone Graph

核心

graph traversal


---

25 Network Delay Time

核心

Dijkstra


---

六、Dynamic Programming（少但可能出现）

26 Climbing Stairs

核心

dp


---

27 House Robber

核心

dp


---

28 Coin Change

核心

dp


---

29 Longest Increasing Subsequence

核心

dp / binary search


---

30 Word Break

核心

dp + hashmap


---

高频程度排序（最重要）

如果时间有限，只练 前15题：

Two Sum
Top K Frequent
Subarray Sum K
Group Anagrams
Longest Substring Without Repeating
Valid Parentheses
Merge Intervals
Meeting Rooms
Minimum Size Subarray Sum
Maximum Sliding Window
Permutation in String
Number of Islands
Course Schedule
Climbing Stairs
House Robber

这些几乎覆盖 80% SHL题型。


---

备考强度建议

如果你准备 2周冲刺

每天练：

5题

总计：

70题

就基本稳过 SHL Code Test。


---

如果你愿意，我可以继续给你一份 更重要的东西：

SHL Code Test 真正的“题型套路表”（10种模板）

很多题其实只是 10种算法模板的变体。
掌握这10个模板，基本能秒掉 80%题目。