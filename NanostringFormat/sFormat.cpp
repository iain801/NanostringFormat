#include "sFormat.h"
#include <algorithm>
#include <iostream>

using namespace libxl;

sFormat::sFormat(std::wstring srcPath, std::wstring sheetLabel)
	: srcPath(srcPath), sheetLabel(sheetLabel)
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

	src->load(srcPath.c_str());
	std::wcout << "Loaded books" << std::endl;

	output = xlCreateXMLBook();
	output->setKey(L"Iain Weissburg", L"windows-2a242a0d01cfe90a6ab8666baft2map2");
}

sFormat::~sFormat() {
	output->save(srcPath.replace(srcPath.find(L".xls"), 4, L"_processed.xls").c_str());
	
	src->release();
	output->release();
}


bool sFormat::isSheets()
{
	return getSheet(src, sheetLabel) != -1;
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
	outSheet = output->addSheet(sheetLabel.c_str());
	
	outSheet->setCol(0,20, 50);

	//TODO: Headers

	int outRow = 1;

	for (int col = 3; col < srcSheet->lastFilledCol(); col++)
	{
		parsedSample parse(srcSheet->readStr(2, col));

		for (int row = 15; row < srcSheet->lastFilledRow(); row++)
		{
			CopyCell(0, outRow, col, 0); // Source File Name
			CopyCell(2, outRow, col, 1); // Sample ID
			CopyCell(4, outRow, col, 6); // Sample Analysis Date 
			CopyCell(6, outRow, col, 7); //Gene RLF

			outSheet->writeNum(outRow, 2, parse.SANum); // Sample Accession Number
			outSheet->writeStr(outRow, 3, parse.patientID.c_str()); // Patient ID
			outSheet->writeStr(outRow, 4, parse.visit.c_str()); // Visit
			outSheet->writeStr(outRow, 5, parse.timepoint.c_str()); // Timepoint
			
			CopyCell(row, outRow, 0, 8); // Gene Classification
			CopyCell(row, outRow, 1, 9); // Gene Name
			CopyCell(row, outRow, 2, 10); // Reference Sequence Number
			CopyCell(row, outRow, col, 11); // Normalized Nanostring Data
			
			outRow++;
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
	Format* srcFormat = srcSheet->cellFormat(srcRow, srcCol);
	if (srcSheet->isFormula(srcRow, srcCol))
	{
		const wchar_t* s = srcSheet->readFormula(srcRow, srcCol);
		outSheet->writeFormula(destRow, destCol, s, srcFormat);
		std::wcout << std::wstring(s ? s : L"null") << " [formula]" << std::endl;
	}
	else
	{
		switch (cellType)
		{
		case CELLTYPE_EMPTY:
		{
			//std::wcout << "[empty]" << std::endl;
			outSheet->writeStr(destRow, destCol, L"", 0L, CELLTYPE_EMPTY);
			break;
		}
		case CELLTYPE_NUMBER:
		{
			double d = srcSheet->readNum(srcRow, srcCol);
			std::wcout << d << " [number] << std::endl";
			
			outSheet->writeNum(destRow, destCol, d, 0L);
			if (srcSheet->isDate(srcRow, srcCol))
			{
				Format* form = output->addFormat();
				form->setNumFormat(NUMFORMAT_DATE);
				outSheet->setCellFormat(destRow, destCol, form);
			}
			break;
		}
		case CELLTYPE_STRING:
		{
			const wchar_t* s = srcSheet->readStr(srcRow, srcCol);
			std::wcout << std::wstring(s ? s : L"null") << " [string]" << std::endl;
			outSheet->writeStr(destRow, destCol, s, 0L);
			break;
		}
		case CELLTYPE_BOOLEAN:
		{
			bool b = srcSheet->readBool(srcRow, srcCol);
			std::wcout << (b ? "true" : "false") << " [boolean]" << std::endl;
			outSheet->writeBool(destRow, destCol, b, 0L);
			break;
		}
		case CELLTYPE_BLANK:
		{
			//std::wcout << "[blank]" << std::endl;
			outSheet->writeBlank(destRow, destCol, 0L);
			break;
		}
		case CELLTYPE_ERROR:
		{
			auto e = srcSheet->readError(srcRow, srcCol);
			std::wcout << "[error]" << std::endl;
			outSheet->writeError(destRow, destCol, e, 0L);
			break;
		}
		}
	}
}

sFormat::parsedSample::parsedSample(std::wstring sampleID)
{
	SANum = stoi(sampleID.substr(0, 8));
	patientID = sampleID.substr(10, 10);
	visit = sampleID.substr(22, 4);
	timepoint = sampleID.substr(28);
}
