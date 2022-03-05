/*
 *
 *  File: docwriter.cpp
 *  Author: isee15
 *  Copyright (c) 2022 isee15
 */

#include "docwriter.h"

#include <algorithm>
#include <iostream>
#include <sstream>
#include <string>

#include "docwriterutil.h"
#include "pugixml/pugixml.hpp"
#include "rapidjson/document.h"
#include "rapidjson/stringbuffer.h"
#include "rapidjson/writer.h"
#include "zip/zip.h"

using namespace rapidjson;

std::shared_ptr<pugi::xml_document> DocWriter::CreateOFDXml() {
    std::shared_ptr<pugi::xml_document> doc = std::make_shared<pugi::xml_document>();

    pugi::xml_node declaration_node = doc->append_child(pugi::node_declaration);
    declaration_node.append_attribute("version") = "1.0";
    declaration_node.append_attribute("encoding") = "UTF-8";

    pugi::xml_node node = doc->append_child("ofd:OFD");
    node.append_attribute("xmlns:ofd") = "http://www.ofdspec.org/2016";
    node.append_attribute("DocType") = "OFD";
    node.append_attribute("Version") = "1.1";

    // add description node with text child
    pugi::xml_node docBody = node.append_child("ofd:DocBody");
    pugi::xml_node docInfo = docBody.append_child("ofd:DocInfo");
    std::string docId = DocWriterUtil::GenerateUuid();
    std::string author = "xicheng";
    std::string creationDate = DocWriterUtil::Today();
    std::string modDate = DocWriterUtil::Today();
    std::string creator = "xiyun";

    docInfo.append_child("ofd:DocID").append_child(pugi::node_pcdata).set_value(docId.c_str());
    docInfo.append_child("ofd:Author").append_child(pugi::node_pcdata).set_value(author.c_str());
    docInfo.append_child("ofd:CreationDate")
            .append_child(pugi::node_pcdata)
            .set_value(creationDate.c_str());
    docInfo.append_child("ofd:ModDate").append_child(pugi::node_pcdata).set_value(modDate.c_str());
    docInfo.append_child("ofd:Creator").append_child(pugi::node_pcdata).set_value(creator.c_str());

    pugi::xml_node docRoot = docBody.append_child("ofd:DocRoot");
    docRoot.append_child(pugi::node_pcdata).set_value("Doc_0/Document.xml");

    // doc->print(std::cout);

    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreateDocumentXml(int pageCount) {
    std::shared_ptr<pugi::xml_document> doc = std::make_shared<pugi::xml_document>();

    pugi::xml_node declaration_node = doc->append_child(pugi::node_declaration);
    declaration_node.append_attribute("version") = "1.0";
    declaration_node.append_attribute("encoding") = "UTF-8";

    pugi::xml_node node = doc->append_child("ofd:Document");
    node.append_attribute("xmlns:ofd") = "http://www.ofdspec.org/2016";

    pugi::xml_node commonData = node.append_child("ofd:CommonData");
    commonData.append_child("ofd:MaxUnitID")
            .append_child(pugi::node_pcdata)
            .set_value(std::to_string(GenerateId()).c_str());
    commonData.append_child("ofd:PageArea")
            .append_child("ofd:PhysicalBox")
            .append_child(pugi::node_pcdata)
            .set_value("0 0 210 297");
    commonData.append_child("ofd:PublicRes")
            .append_child(pugi::node_pcdata)
            .set_value("PublicRes.xml");

    pugi::xml_node pages = node.append_child("ofd:Pages");

    for (int i = 0; i < pageCount; i++) {
        pugi::xml_node page = pages.append_child("ofd:Page");
        page.append_attribute("ID") = std::to_string(GenerateId()).c_str();
        std::string baseLoc = "Pages/Page_";
        baseLoc += std::to_string(i);
        baseLoc += "/Content.xml";
        page.append_attribute("BaseLoc") = baseLoc.c_str();
    }

    // doc->print(std::cout);

    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreatePageXml(Document& jsonDoc) {
    std::shared_ptr<pugi::xml_document> doc = std::make_shared<pugi::xml_document>();

    pugi::xml_node declaration_node = doc->append_child(pugi::node_declaration);
    declaration_node.append_attribute("version") = "1.0";
    declaration_node.append_attribute("encoding") = "UTF-8";

    pugi::xml_node node = doc->append_child("ofd:Page");
    node.append_attribute("xmlns:ofd") = "http://www.ofdspec.org/2016";

    pugi::xml_node content = node.append_child("ofd:Content");
    pugi::xml_node layer = content.append_child("ofd:Layer");
    layer.append_attribute("ID") = std::to_string(GenerateId()).c_str();

    Value lineArr = jsonDoc["doc"].GetArray();
    size_t len = lineArr.Size();
    int imgW = jsonDoc["width"].GetInt();
    int imgH = jsonDoc["height"].GetInt();
    int pageW = 210;
    int pageH = 297;
    float scale = 1;
    if (pageW < imgW || pageH < imgH) {
        scale = std::min(pageW * 1.0f / imgW, pageH * 1.0f / imgH);
    }
    for (size_t i = 0; i < len; i++) {
        const Value& line = lineArr[i];
        std::string strContent = line["text"].GetString();

        pugi::xml_node textObject = layer.append_child("ofd:TextObject");
        textObject.append_attribute("ID") = std::to_string(GenerateId()).c_str();
        float boxW = line["w"].GetFloat();
        float boxH = line["h"].GetFloat();
        std::string boundary = std::to_string(line["x"].GetInt() * scale);
        boundary += " ";
        boundary += std::to_string(line["y"].GetInt() * scale);
        boundary += " ";
        boundary += std::to_string(boxW * scale);
        boundary += " ";
        boundary += std::to_string(boxH * scale);

        float fontSize = std::min(boxH * scale * 0.7, boxW / strContent.size() * 0.8);

        textObject.append_attribute("Boundary") = boundary.c_str();
        textObject.append_attribute("Font") = 1;
        textObject.append_attribute("Size") = std::to_string(fontSize).c_str();
        pugi::xml_node textCode = textObject.append_child("ofd:TextCode");
        textCode.append_attribute("X") = "0.4";
        textCode.append_attribute("Y") = std::to_string(fontSize * 0.88).c_str();
        // 2.6246 2.6668
        std::string deltaX = "2.6246";
        deltaX = std::to_string(fontSize);
        for (int i = 1; i < strContent.size() - 1; ++i) {
            deltaX += " ";
            deltaX += std::to_string(fontSize);
        }

        textCode.append_attribute("DeltaX") = deltaX.c_str();
        textCode.append_child(pugi::node_pcdata).set_value(strContent.c_str());
    }

    // doc->print(std::cout);

    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreateDocumentResXml() {
    std::shared_ptr<pugi::xml_document> doc = std::make_shared<pugi::xml_document>();

    pugi::xml_node declaration_node = doc->append_child(pugi::node_declaration);
    declaration_node.append_attribute("version") = "1.0";
    declaration_node.append_attribute("encoding") = "UTF-8";

    pugi::xml_node node = doc->append_child("ofd:Res");
    node.append_attribute("xmlns:ofd") = "http://www.ofdspec.org/2016";
    node.append_attribute("BaseLoc") = "Res";

    // doc->print(std::cout);

    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreatePublicResXml() {
    std::shared_ptr<pugi::xml_document> doc = std::make_shared<pugi::xml_document>();

    pugi::xml_node declaration_node = doc->append_child(pugi::node_declaration);
    declaration_node.append_attribute("version") = "1.0";
    declaration_node.append_attribute("encoding") = "UTF-8";

    pugi::xml_node node = doc->append_child("ofd:Res");
    node.append_attribute("xmlns:ofd") = "http://www.ofdspec.org/2016";
    node.append_attribute("BaseLoc") = "Res";

    pugi::xml_node fonts = node.append_child("ofd:Fonts");
    pugi::xml_node font = fonts.append_child("ofd:Font");
    font.append_attribute("ID") = "1";
    font.append_attribute("FontName") = "宋体";
    font = fonts.append_child("ofd:Font");
    font.append_attribute("ID") = "2";
    font.append_attribute("FontName") = "Calibri";

    // doc->print(std::cout);

    return doc;
}

int DocWriter::GenerateId() {
    return ++m_id;
}

int DocWriter::GenerateOfd(std::vector<std::string> ocrRet, const char* outputPath) {
    std::string strPdfName = outputPath;

    std::shared_ptr<pugi::xml_document> ofdDoc = CreateOFDXml();
    std::shared_ptr<pugi::xml_document> publicResDoc = CreatePublicResXml();
    std::vector<std::shared_ptr<pugi::xml_document>> pageDocArr;
    for (int p = 0; p < ocrRet.size(); ++p) {
        const char* content = ocrRet[p].c_str();
        Document doc;
        doc.Parse(content);
        std::shared_ptr<pugi::xml_document> page = CreatePageXml(doc);
        pageDocArr.push_back(page);
    }
    std::shared_ptr<pugi::xml_document> docDoc = CreateDocumentXml(ocrRet.size());

    struct zip_t* zip = zip_open(outputPath, ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');
    {
        zip_entry_open(zip, "OFD.xml");
        {
            std::stringstream ss;
            ofdDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        zip_entry_open(zip, "Doc_0/Document.xml");
        {
            std::stringstream ss;
            docDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        zip_entry_open(zip, "Doc_0/PublicRes.xml");
        {
            std::stringstream ss;
            publicResDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        for (int i = 0; i < pageDocArr.size(); ++i) {
            std::string baseLoc = "Doc_0/Pages/Page_";
            baseLoc += std::to_string(i);
            baseLoc += "/Content.xml";
            zip_entry_open(zip, baseLoc.c_str());
            {
                std::stringstream ss;
                pageDocArr[i]->save(ss, "  ");
                std::string buf = ss.str();
                zip_entry_write(zip, buf.c_str(), buf.size());
            }
            zip_entry_close(zip);
        }
    }
    zip_close(zip);

    return 0;
}

// Word generate
std::shared_ptr<pugi::xml_document> DocWriter::CreateWordGlobalRelsXml() {
    std::shared_ptr<pugi::xml_document> doc(new pugi::xml_document);

    pugi::xml_node declNode = doc->prepend_child(pugi::node_declaration);
    declNode.append_attribute("version") = "1.0";
    declNode.append_attribute("encoding") = "UTF-8";

    pugi::xml_node rootNode = doc->append_child("Relationships");
    rootNode.append_attribute("xmlns") =
            "http://schemas.openxmlformats.org/package/2006/relationships";

    pugi::xml_node relNode = rootNode.append_child("Relationship");
    relNode.append_attribute("Id") = "rId1";
    relNode.append_attribute("Type") =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            "officeDocument";
    relNode.append_attribute("Target") = "word/document.xml";

    // doc->print(std::cout);

    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreateWordDocumentXml(
        std::vector<std::string> ocrRet) {
    std::shared_ptr<pugi::xml_document> doc(new pugi::xml_document);

    pugi::xml_node declNode = doc->prepend_child(pugi::node_declaration);
    declNode.append_attribute("version") = "1.0";
    declNode.append_attribute("encoding") = "UTF-8";

    pugi::xml_node rootNode = doc->append_child("w:document");
    rootNode.append_attribute("xmlns:wpc") =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    rootNode.append_attribute("xmlns:mc") =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";
    rootNode.append_attribute("xmlns:o") = "urn:schemas-microsoft-com:office:office";
    rootNode.append_attribute("xmlns:r") =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    rootNode.append_attribute("xmlns:m") =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
    rootNode.append_attribute("xmlns:v") = "urn:schemas-microsoft-com:vml";
    rootNode.append_attribute("xmlns:wp14") =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    rootNode.append_attribute("xmlns:wp") =
            "http://schemas.openxmlformats.org/drawingml/2006/"
            "wordprocessingDrawing";
    rootNode.append_attribute("xmlns:w10") = "urn:schemas-microsoft-com:office:word";
    rootNode.append_attribute("xmlns:w") =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    rootNode.append_attribute("xmlns:w14") = "http://schemas.microsoft.com/office/word/2010/wordml";
    rootNode.append_attribute("xmlns:wpg") =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    rootNode.append_attribute("xmlns:wpi") =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
    rootNode.append_attribute("xmlns:wne") = "http://schemas.microsoft.com/office/word/2006/wordml";
    rootNode.append_attribute("xmlns:wps") =
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    rootNode.append_attribute("mc:Ignorable") = "w14 wp14";

    pugi::xml_node bodyNode = rootNode.append_child("w:body");

    for (int p = 0; p < ocrRet.size(); ++p) {
        pugi::xml_node pNode = bodyNode.append_child("w:p");
        pugi::xml_node pPrNode = pNode.append_child("w:pPr");
        pugi::xml_node pStyleNode = pPrNode.append_child("w:pStyle");
        pStyleNode.append_attribute("w:val") = "NormalWeb";
        pugi::xml_node spacingNode = pPrNode.append_child("w:spacing");
        spacingNode.append_attribute("w:before") = "120";
        spacingNode.append_attribute("w:after") = "120";

        pugi::xml_node jcNode = pPrNode.append_child("w:jc");
        jcNode.append_attribute("w:val") = "left";

        const char* content = ocrRet[p].c_str();
        Document doc;
        doc.Parse(content);
        std::string strContent = doc["text"].GetString();

        std::vector<std::string> strArr = DocWriterUtil::Split(strContent, '\n');
        int fontSize = doc["fontSize"].GetInt();

        pugi::xml_node rNode = pNode.append_child("w:r");

        pugi::xml_node rPrNode = rNode.append_child("w:rPr");

        pugi::xml_node szNode = rPrNode.append_child("w:sz");
        szNode.append_attribute("w:val") = std::to_string(fontSize).c_str();

        for (int i = 0; i < strArr.size(); ++i) {
            std::cout << strArr[i] << std::endl;
            pugi::xml_node tNode = rNode.append_child("w:t");
            tNode.append_attribute("xml:space") = "preserve";
            tNode.append_child(pugi::node_pcdata).set_value(strArr[i].c_str());

            if (strArr.size() > 1 && i != strArr.size() - 1) {
                rNode.append_child("w:br");
            }
        }
    }

    pugi::xml_node sectPrNode = bodyNode.append_child("w:sectPr");
    sectPrNode.append_attribute("w:rsidR") = "005F670F";

    pugi::xml_node pgSzNode = sectPrNode.append_child("w:pgSz");
    pgSzNode.append_attribute("w:w") = "12240";
    pgSzNode.append_attribute("w:h") = "15840";

    pugi::xml_node pgMarNode = sectPrNode.append_child("w:pgMar");
    pgMarNode.append_attribute("w:top") = "1440";
    pgMarNode.append_attribute("w:right") = "1440";
    pgMarNode.append_attribute("w:bottom") = "1440";
    pgMarNode.append_attribute("w:left") = "1440";
    pgMarNode.append_attribute("w:header") = "720";
    pgMarNode.append_attribute("w:footer") = "720";
    pgMarNode.append_attribute("w:gutter") = "0";

    pugi::xml_node colsNode = sectPrNode.append_child("w:cols");
    colsNode.append_attribute("w:space") = "720";

    pugi::xml_node docGridNode = sectPrNode.append_child("w:docGrid");
    docGridNode.append_attribute("w:linePitch") = "360";

    // doc->print(std::cout);
    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreateWordRelsXml() {
    std::shared_ptr<pugi::xml_document> doc(new pugi::xml_document);

    pugi::xml_node declNode = doc->prepend_child(pugi::node_declaration);
    declNode.append_attribute("version") = "1.0";
    declNode.append_attribute("encoding") = "UTF-8";

    pugi::xml_node rootNode = doc->append_child("Relationships");
    rootNode.append_attribute("xmlns") =
            "http://schemas.openxmlformats.org/package/2006/relationships";

    // doc->print(std::cout);
    return doc;
}

std::shared_ptr<pugi::xml_document> DocWriter::CreateWordContentTypesXml() {
    std::shared_ptr<pugi::xml_document> doc(new pugi::xml_document);

    pugi::xml_node declNode = doc->prepend_child(pugi::node_declaration);
    declNode.append_attribute("version") = "1.0";
    declNode.append_attribute("encoding") = "UTF-8";

    pugi::xml_node rootNode = doc->append_child("Types");
    rootNode.append_attribute("xmlns") =
            "http://schemas.openxmlformats.org/package/2006/content-types";

    pugi::xml_node defaultNode = rootNode.append_child("Default");
    defaultNode.append_attribute("Extension") = "rels";
    defaultNode.append_attribute("ContentType") =
            "application/vnd.openxmlformats-package.relationships+xml";

    defaultNode = rootNode.append_child("Default");
    defaultNode.append_attribute("Extension") = "xml";
    defaultNode.append_attribute("ContentType") = "application/xml";

    pugi::xml_node overrideNode = rootNode.append_child("Override");
    overrideNode.append_attribute("PartName") = "/word/document.xml";
    overrideNode.append_attribute("ContentType") =
            "application/"
            "vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    // doc->print(std::cout);
    return doc;
}

int DocWriter::GenerateDocx(std::vector<std::string> ocrRet, const char* outputPath) {
    std::shared_ptr<pugi::xml_document> globalRelsDoc = CreateWordGlobalRelsXml();
    std::shared_ptr<pugi::xml_document> documentDoc = CreateWordDocumentXml(ocrRet);
    std::shared_ptr<pugi::xml_document> wordRelsDoc = CreateWordRelsXml();
    std::shared_ptr<pugi::xml_document> contentTypeDoc = CreateWordContentTypesXml();

    zip_t* zip = zip_open(outputPath, ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');
    {
        zip_entry_open(zip, "[Content_Types].xml");
        {
            std::stringstream ss;
            contentTypeDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        zip_entry_open(zip, "_rels/.rels");
        {
            std::stringstream ss;
            globalRelsDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        zip_entry_open(zip, "word/document.xml");
        {
            std::stringstream ss;
            documentDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);

        zip_entry_open(zip, "word/_rels/document.xml.rels");
        {
            std::stringstream ss;
            wordRelsDoc->save(ss, "  ");
            std::string buf = ss.str();
            zip_entry_write(zip, buf.c_str(), buf.size());
        }
        zip_entry_close(zip);
    }
    zip_close(zip);

    return 0;
}
