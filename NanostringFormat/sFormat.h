#pragma once
#include "libxl.h"
#include <string>
#include <list>

class sFormat
{
public:
	sFormat(std::wstring srcPath, std::wstring sheetLabel);
	~sFormat();

	void CopySheet();
	bool isSheets();

private:
	struct parsedSample {
		//parsed sampleID
		int SANum;
		std::wstring patientID,
			visit,
			timepoint;

		parsedSample(std::wstring sampleID);
	};

	std::wstring srcPath;
	std::wstring destPath;
	libxl::Book* src = nullptr;
	libxl::Sheet* srcSheet = nullptr;
	libxl::Book* output = nullptr;
	libxl::Sheet* outSheet = nullptr;

	std::wstring sheetLabel;

	void CopyCell(int row, int col);
	void CopyCell(int srcRow, int destRow, int srcCol, int destCol);
	int getSheet(libxl::Book* book, std::wstring label);
};

