# Word "发现无法读取的内容" 排错记录

## 现象

清洗元数据后，Word 打开文件弹窗报错：
> Word 在 xxx.docx 中发现无法读取的内容。是否恢复此文档的内容?

点击"是"后 Word 可恢复打开，但说明清洗过程破坏了文件结构。

## 排查过程

### 第一轮：怀疑 ZIP 结构问题
- 怀疑点：`[Content_Types].xml` 压缩方式不对（OPC 规范要求 STORED）
- 结论：不是。python-docx 生成的文件本身就 DEFLATED，且能正常打开。

### 第二轮：怀疑命名空间被丢弃（ElementTree 重序列化）
- 怀疑点：ElementTree 序列化时会丢弃未使用的命名空间声明（`xmlns:xsi`、`xmlns:dcmitype`）
- 修复方式：改用正则表达式直接修改 XML 文本，保留所有原始结构
- 结果：命名空间、换行符 `\r\n`、元素开闭格式全部保留，但 Word **仍然报错**

### 第三轮：隔离测试（关键突破）
- 创建 `test_passthrough.docx`（纯 ZIP 复制，不改 XML）→ **正常** ✓
- 创建 4 个单变量测试文件：

| 测试 | 改动 | 结果 |
|------|------|------|
| A | 仅清空 `dc:creator`（字符串） | 正常 ✓ |
| B | 仅删除 `dcterms:created/modified`（日期） | 正常 ✓ |
| C | 仅改 `revision` 为 1 | 正常 ✓ |
| D | 全部组合 | **报错** ✗ |

**结论**：每个单独操作都正常，但组合起来就报错 → 是更多字段的类型问题。

## 根因

OOXML 的 core.xml / app.xml 中，不同元素有不同数据类型：

| 元素 | 类型 | 清空为 "" | 是否合法 |
|------|------|-----------|----------|
| `dc:creator` | xsd:string | 可以 | ✓ |
| `dc:title` | xsd:string | 可以 | ✓ |
| `dcterms:created` | dcterms:W3CDTF | **不行** | ✗ |
| `dcterms:modified` | dcterms:W3CDTF | **不行** | ✗ |
| `cp:lastPrinted` | xsd:dateTime | **不行** | ✗ |
| `TotalTime` | xsd:int | **不行** | ✗ |

**日期和整数类型的字段，空字符串不符合 schema 约束，Word 的 OOXML 校验器直接拒绝。**

之前 ElementTree 方式虽然没有完全解决，但也恰好因为移除了 `xsi:type` 属性改变了一些行为；而正则方式保留结构更完整，反而让 schema 校验更严格地暴露了这个问题。

## 最终修复

```python
# 字符串字段 → 清空文本（保留元素）
_STRING_FIELDS = [dc:creator, dc:title, ...]

# 日期字段 → 完全删除元素（包括属性）
_DATE_FIELDS = [dcterms:created, dcterms:modified, cp:lastPrinted]

# 整数字段 → 设为 "0"
_APP_FIELDS_TO_ZERO = [TotalTime]

# revision → 设为 "1"
```

## 教训

1. OOXML 不是普通 XML，Office 有严格的 schema 校验
2. 空字符串对字符串类型合法，但对日期/整数类型不合法
3. 排查这类问题时，单变量隔离测试是最有效的定位方法
4. 正则方式的"保留原始结构"虽然是好思路，但也意味着要自己管理每个字段的类型安全
