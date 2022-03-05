# docwriter
docwriter is a free library to generate docx or ofd file.

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
