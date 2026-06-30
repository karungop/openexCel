cmake_minimum_required(VERSION 3.20)
project(openexcel C)

find_package(Python REQUIRED COMPONENTS Interpreter Development.Module)
find_package(EXPAT REQUIRED)

set(OPENEXCEL_SOURCES
    src/_openexcel/module.c
    src/_openexcel/workbook.c
    src/_openexcel/worksheet.c
    src/_openexcel/cell.c
    src/_openexcel/string_table.c
    src/_openexcel/styles.c
    src/_openexcel/reader/reader.c
    src/_openexcel/reader/zip_reader.c
    src/_openexcel/reader/xml_workbook.c
    src/_openexcel/reader/xml_sst.c
    src/_openexcel/reader/xml_styles.c
    src/_openexcel/reader/xml_sheet.c
    src/_openexcel/reader/xml_sheet_rels.c
    src/_openexcel/writer/writer.c
    src/_openexcel/writer/zip_writer.c
    src/_openexcel/writer/xml_gen.c
    src/_openexcel/writer/sheet_writer.c
    src/_openexcel/writer/sst_writer.c
    src/_openexcel/third_party/miniz.c
    src/_openexcel/third_party/miniz_tdef.c
    src/_openexcel/third_party/miniz_tinfl.c
    src/_openexcel/third_party/miniz_zip.c
)

python_add_library(_openexcel MODULE WITH_SOABI ${OPENEXCEL_SOURCES})

target_include_directories(_openexcel PRIVATE
    src/_openexcel
    src/_openexcel/third_party
)

target_link_libraries(_openexcel PRIVATE EXPAT::EXPAT)

target_compile_options(_openexcel PRIVATE
    -O3
    -Wall
    -Wextra
    -Wno-unused-parameter
    -Wno-cast-function-type
    -Wno-missing-field-initializers
    $<$<C_COMPILER_ID:GNU,Clang>:-fvisibility=hidden>
)

install(TARGETS _openexcel DESTINATION openexcel)

