# The Flutter tooling requires that developers have CMake 3.10 or later
# installed. You should not increase this version, as doing so will cause
# the plugin to fail to compile for some customers of the plugin.
cmake_minimum_required(VERSION 3.10)

project(dartxlsxwriter_library VERSION 0.0.1 LANGUAGES C)

add_library(dartxlsxwriter SHARED
  "src/app.c"
  "src/chart.c"
  "src/chartsheet.c"
  "src/comment.c"
  "src/content_types.c"
  "src/core.c"
  "src/custom.c"
  "src/drawing.c"
  "src/format.c"
  "src/hash_table.c"
  "src/metadata.c"
  "src/packager.c"
  "src/relationships.c"
  "src/shared_strings.c"
  "src/styles.c"
  "src/table.c"
  "src/theme.c"
  "src/utility.c"
  "src/vml.c"
  "src/workbook.c"
  "src/worksheet.c"
  "src/xmlwriter.c"
  "third_party/minizip/minizip.c"
  "third_party/minizip/zip.c"
  "third_party/minizip/ioapi.c"
  "third_party/tmpfileplus/tmpfileplus.c"
)

set_target_properties(dartxlsxwriter PROPERTIES
  PUBLIC_HEADER xlsxwriter.h
  OUTPUT_NAME "dartxlsxwriter"
)

if (WIN32)
set_target_properties(ffigen_app PROPERTIES
  WINDOWS_EXPORT_ALL_SYMBOLS ON
)
endif (WIN32)


target_compile_definitions(dartxlsxwriter PUBLIC DART_SHARED_LIB)
