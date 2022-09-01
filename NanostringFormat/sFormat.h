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
		std::wstring sampleID;

		parsedSample(std::wstring sampleID);
	};

	std::wstring srcPath;
	std::wstring destPath;
	libxl::Book* src = nullptr;
	libxl::Sheet* srcSheet = nullptr;
	libxl::Book* output = nullptr;
	libxl::Sheet* outSheet = nullptr;

	libxl::Format* header = nullptr;
	libxl::Format* normal = nullptr;
	libxl::Format* number = nullptr;
	libxl::Format* date = nullptr;

	std::wstring sheetLabel;

	void CopyCell(int row, int col);
	void CopyCell(int srcRow, int destRow, int srcCol, int destCol);
	int getSheet(libxl::Book* book, std::wstring label);

	void initFormats();
};

