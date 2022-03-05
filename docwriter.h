/*
 *  Under MIT license
 *  Author: isee15
 *  Date: Tue Feb 22 2022
 *  docwriter is a free library to generate docx or ofd file.
 *
 */

#ifndef DOC_WRITER_H
#define DOC_WRITER_H
#include <vector>

#include "pugixml/pugixml.hpp"
#include "rapidjson/document.h"

class DocWriter {
public:
    int GenerateOfd(std::vector<std::string> ocrRet, const char* outputPath);
    int GenerateDocx(std::vector<std::string> ocrRet, const char* outputPath);

private:
    // ofd
    std::shared_ptr<pugi::xml_document> CreateOFDXml();
    std::shared_ptr<pugi::xml_document> CreateDocumentXml(int pageCount);
    std::shared_ptr<pugi::xml_document> CreatePageXml(rapidjson::Document& lineArr);
    std::shared_ptr<pugi::xml_document> CreateDocumentResXml();
    std::shared_ptr<pugi::xml_document> CreatePublicResXml();

    // Word docx
    std::shared_ptr<pugi::xml_document> CreateWordGlobalRelsXml();
    std::shared_ptr<pugi::xml_document> CreateWordDocumentXml(std::vector<std::string> ocrRet);
    std::shared_ptr<pugi::xml_document> CreateWordRelsXml();
    std::shared_ptr<pugi::xml_document> CreateWordContentTypesXml();

    int m_id = 2;
    int GenerateId();
};

#endif