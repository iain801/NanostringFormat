#include "cFrame.h"
#include <string>
#include <iostream>

wxBEGIN_EVENT_TABLE(cFrame, wxFrame)
EVT_BUTTON(10001, PerformTransfer)
EVT_FILEPICKER_CHANGED(10002, ResetButton)
EVT_FILEPICKER_CHANGED(10003, ResetButton)
wxEND_EVENT_TABLE()



//if using output, change to wxSize(340, 400)
cFrame::cFrame() : wxFrame(nullptr, wxID_ANY, "Nanostring Data Format - Erasca", wxPoint(100, 100), wxSize(340, 130), wxDEFAULT_FRAME_STYLE & ~(wxRESIZE_BORDER | wxMAXIMIZE_BOX))
{
	btn1 = new wxButton(this, 10001, "Format", wxPoint(10, 50), wxSize(150, 30));
	labelText = new wxStaticText(this, wxID_ANY, "Target Sheet: ", wxPoint(165, 57));
	labelInput = new wxTextCtrl(this, wxID_ANY, "Normalized Data", wxPoint(235, 55), wxSize(75, 20));
	srcText = new wxStaticText(this, wxID_ANY, "Input: ", wxPoint(10, 2));
	srcFile = new wxFilePickerCtrl(this, 10002, "", "", "XLSX and XLS files (*.xlsx;*.xls)|*.xlsx;*.xls", wxPoint(10, 20), wxSize(300, 20));
	//dstText = new wxStaticText(this, wxID_ANY, "Output: ", wxPoint(10, 42));
	//dstFile = new wxFilePickerCtrl(this, 10003, "", "", "XLSX and XLS files (*.xlsx;*.xls)|*.xlsx;*.xls", wxPoint(10, 60), wxSize(300, 20));
	//output = new wxTextCtrl (this, wxID_ANY, "", wxPoint(10, 135), wxSize(300, 200), wxTE_READONLY + wxTE_MULTILINE);

	//wxStreamToTextRedirector redirect(output);

	//std::cout << "ENSURE DATA IS ORDERED IN ALL SHEETS" << std::endl;
}

cFrame::~cFrame()
{

}

void cFrame::PerformTransfer(wxCommandEvent& evt)
{
	//wxStreamToTextRedirector redirect(output, std::wcout);
	btn1->SetLabelText("Running...");
	std::wstring srcPath = srcFile->GetPath().ToStdWstring();
	//std::wstring destPath = dstFile->GetPath().ToStdWstring();
	auto labelText = labelInput->GetLineText(0).ToStdString();
	if (srcPath.find(L".xls") == std::wstring::npos)
		btn1->SetLabelText("Invalid Input");
	//else if (destPath.find(L".xls") == std::wstring::npos)
		//btn1->SetLabelText("Invalid Output");
	//output->Clear();
	else if (!labelText.empty())
	{
		std::wstring label = labelInput->GetLineText(0).ToStdWstring();
		sFormat* update = new sFormat(srcPath, label);
		if (update->isSheets())
		{
			update->CopySheet();
			btn1->SetLabelText("Finished!");
		}
		else
		{
			btn1->SetLabelText("Target Not Found");
		}
		delete update;
	}
	else
	{
		btn1->SetLabelText("No Sheet Selected");
	}
	evt.Skip();
}

void cFrame::ResetButton(wxFileDirPickerEvent& evt)
{
	btn1->SetLabelText("Format");
	evt.Skip();
}