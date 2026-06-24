#pragma once
#include "../string_table.h"
#include "xml_gen.h"

/* Emit sharedStrings.xml XML into b from the given string table. */
void oxl_write_sst(OxlXmlBuf *b, const OxlStringTable *sst);
