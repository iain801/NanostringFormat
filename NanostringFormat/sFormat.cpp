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

	initFormats();
}

sFormat::~sFormat() {
	output->save(srcPath.replace(srcPath.find(L".xls"), 4, L"_processed.xls").c_str());
	
	src->release();
	output->release();
}


void sFormat::initFormats()
{
	header = output->addFormat();
	header->setFillPattern(FILLPATTERN_SOLID);
	header->setPatternForegroundColor(COLOR_ICEBLUE_CF);
	header->setBorderLeft(BORDERSTYLE_MEDIUM);
	header->setBorderRight(BORDERSTYLE_MEDIUM);
	header->setBorderTop(BORDERSTYLE_MEDIUM);
	header->setBorderBottom(BORDERSTYLE_MEDIUM);
	header->setAlignH(ALIGNH_LEFT);
	header->setWrap(true);

	normal = output->addFormat();
	normal->setAlignH(ALIGNH_LEFT);
	normal->setWrap(true);
	normal->setNumFormat(NUMFORMAT_GENERAL);

	number = output->addFormat();
	number->setAlignH(ALIGNH_LEFT);
	number->setWrap(true);
	number->setNumFormat(NUMFORMAT_NUMBER);

	date = output->addFormat();
	date->setAlignH(ALIGNH_LEFT);
	date->setWrap(true);
	date->setNumFormat(NUMFORMAT_DATE);
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
	
	outSheet->setCol(0,1, 40);
	outSheet->setCol(2, 6, 15);
	outSheet->setCol(7, 7, 22);
	outSheet->setCol(8, 8, 18);
	outSheet->setCol(9, 9, 15);
	outSheet->setCol(10, 11, 18);

	outSheet->writeStr(0, 0, L"Source File Name", header);
	outSheet->writeStr(0, 1, L"Sample ID", header);
	outSheet->writeStr(0, 2, L"Sample Accession Number", header);
	outSheet->writeStr(0, 3, L"Patient ID", header);
	outSheet->writeStr(0, 4, L"Visit", header);
	outSheet->writeStr(0, 5, L"Timepoint", header);
	outSheet->writeStr(0, 6, L"Sample Analysis Date", header);
	outSheet->writeStr(0, 7, L"Gene RLF", header);
	outSheet->writeStr(0, 8, L"Gene Classification", header);
	outSheet->writeStr(0, 9, L"Gene Name", header);
	outSheet->writeStr(0, 10, L"RefSeq Number", header);
	outSheet->writeStr(0, 11, L"Normalized Nanostring Data", header);

	int outRow = 1;
	for (int col = 3; col < srcSheet->lastFilledCol(); col++)
	{
		parsedSample parse(srcSheet->readStr(2, col));
		for (int row = 15; row < srcSheet->lastFilledRow(); row++)
		{
			CopyCell(0, outRow, col, 0);				// Source File Name

			outSheet->writeStr(outRow, 1, parse.sampleID.c_str(), normal);		// Sample ID
			outSheet->writeNum(outRow, 2, parse.SANum, number);				// Sample Accession Number
			outSheet->writeStr(outRow, 3, parse.patientID.c_str(), normal);	// Patient ID
			outSheet->writeStr(outRow, 4, parse.visit.c_str(), normal);		// Visit
			outSheet->writeStr(outRow, 5, parse.timepoint.c_str(), normal);	// Timepoint

			CopyCell(4, outRow, col, 6);				// Sample Analysis Date 
			CopyCell(6, outRow, col, 7);				// Gene RLF
			CopyCell(row, outRow, 0, 8);				// Gene Classification
			CopyCell(row, outRow, 1, 9);				// Gene Name
			CopyCell(row, outRow, 2, 10);			// Reference Sequence Number
			CopyCell(row, outRow, col, 11);			// Normalized Nanostring Data
			
			outRow++;
		}
	}

	//TODO: Column sorting
}

void sFormat::CopyCell(int row, int col)
{
	CopyCell(row, row, col, col);
}

void sFormat::CopyCell(int srcRow, int destRow, int srcCol, int destCol)
{
	auto cellType = srcSheet->cellType(srcRow, srcCol);

	if (srcSheet->isFormula(srcRow, srcCol))
	{
		const wchar_t* s = srcSheet->readFormula(srcRow, srcCol);
		outSheet->writeFormula(destRow, destCol, s, normal);
		std::wcout << std::wstring(s ? s : L"null") << " [formula]" << std::endl;
	}
	else
	{
		switch (cellType)
		{
		case CELLTYPE_EMPTY:
		{
			//std::wcout << "[empty]" << std::endl;
			outSheet->writeStr(destRow, destCol, L"", normal, CELLTYPE_EMPTY);
			break;
		}
		case CELLTYPE_NUMBER:
		{
			double d = srcSheet->readNum(srcRow, srcCol);
			std::wcout << d << " [number] << std::endl";
			if (srcSheet->isDate(srcRow, srcCol))
				outSheet->writeNum(destRow, destCol, d, date);
			else
				outSheet->writeNum(destRow, destCol, d, number);
			break;
		}
		case CELLTYPE_STRING:
		{
			const wchar_t* s = srcSheet->readStr(srcRow, srcCol);
			std::wcout << std::wstring(s ? s : L"null") << " [string]" << std::endl;
			outSheet->writeStr(destRow, destCol, s, normal);
			break;
		}
		case CELLTYPE_BOOLEAN:
		{
			bool b = srcSheet->readBool(srcRow, srcCol);
			std::wcout << (b ? "true" : "false") << " [boolean]" << std::endl;
			outSheet->writeBool(destRow, destCol, b, normal);
			break;
		}
		case CELLTYPE_BLANK:
		{
			//std::wcout << "[blank]" << std::endl;
			outSheet->writeBlank(destRow, destCol, normal);
			break;
		}
		case CELLTYPE_ERROR:
		{
			auto e = srcSheet->readError(srcRow, srcCol);
			std::wcout << "[error]" << std::endl;
			outSheet->writeError(destRow, destCol, e, normal);
			break;
		}
		}
	}
}

sFormat::parsedSample::parsedSample(std::wstring SID) :
	sampleID(SID)
{
	
	while (sampleID.find(L"  ") != std::wstring::npos)
		sampleID = sampleID.replace(sampleID.find(L"  "), 2, L" ");
	
	std::size_t i = 0;
	SANum = stoi(sampleID.substr(0, sampleID.find(L" ")), &i);
	patientID = sampleID.substr(++i, 9);
	i += 10;
	int l = sampleID.find(L" ", i) - i;
	visit = sampleID.substr(i, l);
	i += l;
	timepoint = sampleID.substr(++i);
}
