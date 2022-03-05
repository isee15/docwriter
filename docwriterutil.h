/*
 *  Under MIT license
 *  Author: isee15
 *  Date: Tue Feb 22 2022
 *  docwriter is a free library to generate docx or ofd file.
 *
 */

#ifndef DOC_WRITER_UTIL_H
#define DOC_WRITER_UTIL_H

#include <iostream>
#include <string>
#include <vector>

        class DocWriterUtil {
public:
    static std::string GenerateUuid();
    static std::string Today();
    static std::vector<std::string> Split(const std::string& s, char delim);
};

#endif // DOC_WRITER_UTIL_H