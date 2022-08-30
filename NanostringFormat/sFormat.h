#pragma once
#include "libxl.h"
#include <string>
#include <list>

class sFormat
{
public:
	sFormat(std::wstring srcPath, std::wstring destPath,
		std::wstring sheetLabel);
	~sFormat();

	void CopySheet();
	bool isSheets();

private:
	std::wstring srcPath;
	std::wstring destPath;
	libxl::Book* src = nullptr;
	libxl::Book* dest = nullptr;
	libxl::Sheet* srcSheet = nullptr;
	libxl::Sheet* destSheet = nullptr;

	std::wstring sheetLabel;

	void CopyCell(int row, int col);
	void CopyCell(int srcRow, int destRow, int srcCol, int destCol);
	int getSheet(libxl::Book* book, std::wstring label);
};

