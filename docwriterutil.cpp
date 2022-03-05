/*
 *
 *  File: docwriterutil.cpp
 *  Author: isee15
 *  Copyright (c) 2022 isee15
 *
 */

#include "docwriterutil.h"

#include <time.h>

#include <sstream>

#ifdef WIN32
#include <objbase.h>
#else
#include <uuid/uuid.h>
#endif

std::string DocWriterUtil::GenerateUuid() {
#ifdef WIN32
    GUID guid;
    CoCreateGuid(&guid);
    char guid_string[40];
    sprintf(guid_string, "%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X", guid.Data1, guid.Data2,
            guid.Data3, guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3], guid.Data4[4],
            guid.Data4[5], guid.Data4[6], guid.Data4[7]);
    return std::string(guid_string);
#else
    uuid_t uuid;
    uuid_generate(uuid);
    char uuid_string[40];
    uuid_unparse(uuid, uuid_string);
    return std::string(uuid_string);
#endif
}

std::string DocWriterUtil::Today() {
    time_t t = time(0);
    struct tm* lt = localtime(&t);
    char buf[32];
    sprintf(buf, "%04d-%02d-%02d", lt->tm_year + 1900, lt->tm_mon + 1, lt->tm_mday);
    return buf;
}

std::vector<std::string>& split(const std::string& s, char delim, std::vector<std::string>& elems) {
    std::stringstream ss(s);
    std::string item;
    while (std::getline(ss, item, delim)) {
        if (item.length() > 0) {
            elems.push_back(item);
        }
    }
    return elems;
}

std::vector<std::string> DocWriterUtil::Split(const std::string& s, char delim) {
    std::vector<std::string> elems;
    split(s, delim, elems);
    return elems;
}