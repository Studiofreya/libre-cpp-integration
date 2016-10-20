#pragma comment(lib, "Ws2_32.lib")
#pragma comment(lib, "isal.lib")
#pragma comment(lib, "icppu.lib")
#pragma comment(lib, "icppuhelper.lib")

// Simple client application using the UnoUrlResolver service.
#include <iostream>

#include <cppuhelper/bootstrap.hxx>
#include <com/sun/star/bridge/XUnoUrlResolver.hpp>
#include <com/sun/star/frame/Desktop.hpp>
#include <com/sun/star/registry/XSimpleRegistry.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>


using namespace com::sun::star::uno;
using namespace com::sun::star::lang;
using namespace com::sun::star::frame;
using namespace com::sun::star::sheet;
using namespace com::sun::star::table;

using ::rtl::OUString;


int insertIntoCell(int CellX, int CellY,
	const std::string & value,
	Reference<XSpreadsheet> sheet)
{
	try {
		Reference<XCell> xCell = nullptr;
		xCell = sheet->getCellByPosition(CellX, CellY);
		xCell->setFormula(OUString::createFromAscii(value.c_str()));
	}
	catch (const com::sun::star::lang::IndexOutOfBoundsException & ex) {
		std::cerr << "Could not get Cell";
		return 1;
	}

	return 0;
}

int insertIntoCell(int CellX, int CellY,
	double value,
	Reference<XSpreadsheet> sheet)
{
	try {
		Reference<XCell> xCell = nullptr;
		xCell = sheet->getCellByPosition(CellX, CellY);
		xCell->setValue(value);
	}
	catch (const com::sun::star::lang::IndexOutOfBoundsException & ex) {
		std::cerr << "Could not get Cell";
		return 1;
	}

	return 0;
}


int start_libreoffice_calc()
{
	try
	{
		// Get the remote office component context
		Reference<XComponentContext> xContext(::cppu::bootstrap());
		if (!xContext.is())
		{
			std::cerr << "No component context!\n";
			return 1;
		}

		// get the remote office service manager
 		Reference<XMultiComponentFactory> xServiceManager(
 			xContext->getServiceManager());
 		if (!xServiceManager.is())
 		{
 			std::cerr << "No service manager!\n";
 			return 1;
 		}

		// get an instance of the remote office desktop UNO service
		// and query the XComponentLoader interface
		Reference<XDesktop2> xComponentLoader = Desktop::create(xContext);

		// open a spreadsheet document
		Reference<XComponent> xComponent(xComponentLoader->loadComponentFromURL(
			OUString("private:factory/scalc"),
			OUString("_blank"), 0,
			Sequence < ::com::sun::star::beans::PropertyValue >()));
		if (!xComponent.is())
		{
			std::cerr << "opening spreadsheet document failed!\n";
			return 1;
		}

		// Reference to the document
		Reference <XSpreadsheetDocument> spreadDoc(xComponent, UNO_QUERY);

		// Get the first sheet
		Reference<XSpreadsheets> rSheets = spreadDoc->getSheets();
		Any rSheet = rSheets->getByName(OUString::createFromAscii("Sheet1"));
		Reference<XSpreadsheet> sheet(rSheet, UNO_QUERY);

		// Insert cells
		insertIntoCell(1, 1, "I was inserted from C++", sheet);
		insertIntoCell(1, 2, 3.14, sheet);
		insertIntoCell(0, 3, "Formula pi^2:", sheet);
		insertIntoCell(1, 3, "=B3*B3", sheet);

	}
	catch (::cppu::BootstrapException & e)
	{
		std::cerr << "Caught BootstrapException: " 
			<< OUStringToOString(e.getMessage(), RTL_TEXTENCODING_ASCII_US).getStr() 
			<< std::endl;
		return 1;
	}
	catch (Exception & e)
	{
		std::cerr << "Caught UNO exception: "
			<< OUStringToOString(e.Message, RTL_TEXTENCODING_ASCII_US).getStr()
			<< std::endl;
		return 1;
	}

	return 0;
}


int main()
{
	start_libreoffice_calc();

	return 0;
}