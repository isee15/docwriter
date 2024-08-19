# docwriter
docwriter is a free library to generate docx or ofd file.

# 全文总结
这段文本主要介绍了一个名为 docwriter 的免费库，用于生成 docx 或 ofd 文件。

**重要亮点**
- **docwriter 库的功能**：能够生成 docx 和 ofd 格式的文件。
- **数据处理**：通过 std::vector<std::string> 类型的 ocrRet 和 ofdRet 来存储相关数据。
- **内容示例**：ocrRet 中包含了如“Old MacDonald Had a Farm\\nRock Scissors Paper\\nClean Up”和“春江潮水连海平”等文本，ofdRet 中包含了如“江南春尽离肠断 20220222 星期二 正月廿二”等内容。
- **生成文件操作**：分别使用 GenerateDocx 函数将 ocrRet 中的数据生成到“./demo.docx”文件，使用 GenerateOfd 函数将 ofdRet 中的数据生成到“./demo.ofd”文件，并提供了 ofd 预览的链接。


# Usage
```
DocWriter docWriter;
std::vector<std::string> ocrRet;
ocrRet.push_back(
        "{\"text\":\"Old MacDonald Had a Farm\\nRock Scissors Paper\\nClean "
        "Up\",\"fontSize\":12}");
ocrRet.push_back("{\"text\":\"春江潮水连海平\",\"fontSize\":12}");
docWriter.GenerateDocx(ocrRet, "./demo.docx");

// ofd 预览 https://51shouzu.xyz/ofd/
std::vector<std::string> ofdRet;
ofdRet.push_back(
        "{\"width\":800,\"height\":600,\"doc\":[{\"text\":"
        "\"江南春尽离肠断20220222星期二正月廿二\",\"x\":12,\"y\":"
        "30,\"w\":400,\"h\":20}]}");
docWriter.GenerateOfd(ofdRet, "./demo.ofd");
```
