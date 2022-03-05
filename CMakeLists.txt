# cmake最低版本号要求
cmake_minimum_required(VERSION 3.10)

# 设置PROJECT_NAME变量
set(PROJECT_NAME DocWriter)

# 指定为C++11 版本
set(CMAKE_CXX_STANDARD 11)

# 设置指定的C++编译器版本是必须的，如果不设置，或者为OFF，则指定版本不可用时，会使用上一版本。
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 设置工程名
project (${PROJECT_NAME})
# 查找当前目录下的所有源文件并存入DIR_SRCS变量
set(HEADERS "${CMAKE_CURRENT_SOURCE_DIR}/docwriter.h"
            "${CMAKE_CURRENT_SOURCE_DIR}/docwriterutil.h")
set(SOURCES docwriter.cpp
            docwriterutil.cpp)

set(THIRD_PARTY_SRC "${CMAKE_CURRENT_SOURCE_DIR}/3rd/zip/zip.c"
"${CMAKE_CURRENT_SOURCE_DIR}/3rd/pugixml/pugixml.cpp")

include_directories("${CMAKE_CURRENT_SOURCE_DIR}/include"
                    "${CMAKE_CURRENT_SOURCE_DIR}/3rd")


add_library(doc_writer STATIC ${SOURCES} ${THIRD_PARTY_SRC})
# 为指定项目添加 include 路径
target_include_directories(doc_writer PUBLIC
                                        3rd
)

# 添加一个可编译的目标到工程
# add_executable (${PROJECT_NAME} ${SOURCES})

set(DEMO_SOURCES demo.cpp)
add_executable(demo ${DEMO_SOURCES} ${THIRD_PARTY_SRC})
target_link_libraries(demo doc_writer)
# 为指定项目添加 include 路径
target_include_directories(demo PUBLIC
                                3rd
)


