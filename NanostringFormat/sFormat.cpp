#include "sFormat.h"
#include <algorithm>
#include <iostream>

using namespace libxl;

sFormat::sFormat(std::wstring srcPath, std::wstring destPath, std::wstring sheetLabel)
	: srcPath(srcPath), destPath(destPath), sheetLabel(sheetLabel)
{
	if (srcPath.compare(destPath) == 0) {
		std::wcout << "ERROR: Duplicate input paths" << std::endl;
		return;
	}

	if (srcPath.find(L".xlsx") != std::wstring::npos)
		src = xlCreateXMLBook();
	else if (srcPath.find(L".xls") != std::wstring::npos)
		src = xlCreateBook();
	else {
		std::wcout << "ERROR: Invaid source filetype" << std::endl;
		return;
	}
	src->setKey(L"Iain Weissburg", L"windows-2a242a0d01cfe90a6ab8666baft2map2");
	std::wcout << "Created source book" << std::endl;

	if (destPath.find(L".xlsx") != std::wstring::npos)
		dest = xlCreateXMLBook();
	else if (destPath.find(L".xls") != std::wstring::npos)
		dest = xlCreateBook();
	else {
		std::wcout << "ERROR: Invaid destination filetype" << std::endl;
		return;
	}
	dest->setKey(L"Iain Weissburg", L"windows-2a242a0d01cfe90a6ab8666baft2map2");
	std::wcout << "Created destination book" << std::endl;

	src->load(srcPath.c_str());
	dest->load(destPath.c_str());
	std::wcout << "Loaded books" << std::endl;
}

sFormat::~sFormat() {
	dest->save(destPath.replace(destPath.find(L".xls"), 4, L"_processed.xls").c_str());
	std::wcout << "Output saved as: " << destPath << std::endl;

	src->release();
	dest->release();
}


bool sFormat::isSheets()
{
	return (getSheet(src, sheetLabel) != -1)
		&& (getSheet(dest, sheetLabel) != -1);
}

int sFormat::getSheet(libxl::Book* book, std::wstring label)
{
	for (int sheet = 0; sheet < book->sheetCount(); sheet++)
	{
		std::wstring sheetName(book->getSheetName(sheet));
		if (sheetName.compare(label) == 0)
			return sheet;
	}
	return -1;
}

void sFormat::CopySheet()
{
	srcSheet = src->getSheet(getSheet(src, sheetLabel));
	int destSheetNum = getSheet(dest, sheetLabel);
	dest->delSheet(destSheetNum);
	destSheet = dest->insertSheet(destSheetNum, srcSheet->name());
	for (int col = 0; col < srcSheet->lastCol(); col++)
	{
		destSheet->setCol(col, col, srcSheet->colWidth(col));
		for (int row = 0; row < srcSheet->lastRow(); row++)
		{
			CopyCell(row, col);
		}
	}
}

void sFormat::CopyCell(int row, int col)
{
	CopyCell(row, row, col, col);
}

void sFormat::CopyCell(int srcRow, int destRow, int srcCol, int destCol)
{
	auto cellType = srcSheet->cellType(srcRow, srcCol);
	auto srcFormat = srcSheet->cellFormat(srcRow, srcCol);
	if (srcSheet->isFormula(srcRow, srcCol))
	{
		const wchar_t* s = srcSheet->readFormula(srcRow, srcCol);
		destSheet->writeFormula(destRow, destCol, s, srcFormat);
		std::wcout << std::wstring(s ? s : L"null") << " [formula]" << std::endl;
	}
	else
	{
		switch (cellType)
		{
		case CELLTYPE_EMPTY:
		{
			//std::wcout << "[empty]" << std::endl;
			destSheet->writeStr(destRow, destCol, L"", srcFormat, CELLTYPE_EMPTY);
			break;
		}
		case CELLTYPE_NUMBER:
		{
			double d = srcSheet->readNum(srcRow, srcCol);
			std::wcout << d << " [number] << std::endl";
			destSheet->writeNum(destRow, destCol, d, srcFormat);
			break;
		}
		case CELLTYPE_STRING:
		{
			const wchar_t* s = srcSheet->readStr(srcRow, srcCol);
			std::wcout << std::wstring(s ? s : L"null") << " [string]" << std::endl;
			destSheet->writeStr(destRow, destCol, s, srcFormat);
			break;
		}
		case CELLTYPE_BOOLEAN:
		{
			bool b = srcSheet->readBool(srcRow, srcCol);
			std::wcout << (b ? "true" : "false") << " [boolean]" << std::endl;
			destSheet->writeBool(destRow, destCol, b, srcFormat);
			break;
		}
		case CELLTYPE_BLANK:
		{
			//std::wcout << "[blank]" << std::endl;
			destSheet->writeBlank(destRow, destCol, srcFormat);
			break;
		}
		case CELLTYPE_ERROR:
		{
			auto e = srcSheet->readError(srcRow, srcCol);
			std::wcout << "[error]" << std::endl;
			destSheet->writeError(destRow, destCol, e, srcFormat);
			break;
		}
		}
	}
}
