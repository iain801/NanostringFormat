#pragma once

#include "wx/wx.h"
#include "wx/filepicker.h"
#include "wx/valnum.h"
#include "sFormat.h"

class cFrame : public wxFrame
{
public:
	cFrame();
	~cFrame();

	wxButton* btn1 = nullptr;
	wxFilePickerCtrl* srcFile = nullptr;
	wxFilePickerCtrl* dstFile = nullptr;
	wxTextCtrl* labelInput = nullptr;
	//wxTextCtrl* output = nullptr;

	wxStaticText* srcText = nullptr;
	wxStaticText* dstText = nullptr;
	wxStaticText* labelText = nullptr;

	void PerformTransfer(wxCommandEvent& evt);
	void ResetButton(wxFileDirPickerEvent& evt);

	wxDECLARE_EVENT_TABLE();

};

