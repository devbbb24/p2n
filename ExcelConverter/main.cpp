#include <iostream>
#include <string>
#include <windows.h>
#include <filesystem>
#include <regex>

#include "OpenXLSX/OpenXLSX.hpp"
#include "nlohmann/json.hpp"

using namespace std;
using namespace nlohmann;

void Log(const string& str) { cout << str << "\n"; }

struct ScriptLine
{
	string type;
	string speaker;
	string text;
	bool italic;
	bool bold;
	int size;
};

std::string getCellStr(OpenXLSX::XLWorksheet& sheet, int row, int col)
{
	auto cell = sheet.cell(row, col);
	if (cell.value().type() == OpenXLSX::XLValueType::Empty)
		return "";
	return cell.value().get<std::string>();
}
int getCellInt(OpenXLSX::XLWorksheet& sheet, int row, int col)
{
	auto cell = sheet.cell(row, col);
	if (cell.value().type() == OpenXLSX::XLValueType::Empty)
		return 0;
	return cell.value().get<int>();
}


regex red("<red>");
regex orange("<org>");
regex blue("<blue>");
regex voice("<vc>");
regex endColor("</red>|</org>|</blue>|</vc>");
regex sizeStart("<fs=");
regex sizeEnd("</f>");

const int COL_TYPE = 1;
const int COL_SPEAKER = 2;
const int COL_TEXT = 3;
const int COL_ITALIC = 4;
const int COL_BOLD = 5;
const int COL_SIZE = 6;

string change_color_style(const string& str)
{
	auto text = str;
	text = std::regex_replace(text, red, string("[color=#ff2222]"));
	text = std::regex_replace(text, orange, string("[color=#E48F5D]"));
	text = std::regex_replace(text, blue, string("[color=#7684A2]"));
	text = std::regex_replace(text, voice, string("[color=#6D63E5]"));
	text = std::regex_replace(text, endColor, string("[/color]"));
	return text;
}
string change_color_html(const string& str)
{
	auto text = str;
	text = std::regex_replace(text, red, string("<font color=#ff2222>"));
	text = std::regex_replace(text, orange, string("<font color=#E48F5D>"));
	text = std::regex_replace(text, blue, string("<font color=#7684A2>"));
	text = std::regex_replace(text, voice, string("<font color=#6D63E5>"));
	text = std::regex_replace(text, endColor, string("</font>"));
	return text;
}

void set_text_style(json& j, OpenXLSX::XLWorksheet& sheet, int row)
{
	j["text"] = change_color_style(getCellStr(sheet, row, COL_TEXT));

	if (getCellInt(sheet, row, COL_ITALIC) == 1)
		j["italic"] = true;
	if (getCellInt(sheet, row, COL_BOLD) == 1)
		j["bold"] = true;
	if (auto size = getCellInt(sheet, row, COL_SIZE))
		j["fontSize"] = size;
}

string set_html(string& text, OpenXLSX::XLWorksheet& sheet, int row)
{
	text = change_color_style(text);

	if (getCellInt(sheet, row, COL_ITALIC) == 1)
		text = format("<i>{}</i>", text);
	if (getCellInt(sheet, row, COL_BOLD) == 1)
		text = format("<b>{}</b>", text);
	if (auto size = getCellInt(sheet, row, COL_SIZE))
		text = format("<font size={}>{}</font>", size, text);
	return text;
}

void set_line(json& j, OpenXLSX::XLWorksheet& sheet, int row)
{
	auto text = getCellStr(sheet, row, COL_TEXT);
	if (!text.empty())
		set_text_style(j, sheet, row);
	else
		j["text"] = "";

	auto speaker = getCellStr(sheet, row, COL_SPEAKER);
	if (!speaker.empty())
		j["speaker"] = change_color_html(speaker);
}

json convert_json(OpenXLSX::XLWorksheet& sheet)
{
	json root;
	root["lines"] = json::array();

	int rowCount = sheet.rowCount();
	for (int i = 2; i <= rowCount; ++i)
	{
		string type = getCellStr(sheet, i, COL_TYPE);
		if (type.empty())
			type = "dialogue";

		if (type == "callout")
		{
			json group;
			group["type"] = "callout";
			group["lines"] = json::array();

			while (i <= rowCount && getCellStr(sheet, i, COL_TYPE) == "callout")
			{
				json line;
				set_line(line, sheet, i);
				group["lines"].push_back(line);

				if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "callout")
					i++;
				else
					break;
			}
			root["lines"].push_back(group);
		}
		else if (type == "choice")
		{
			json choiceGroup;
			choiceGroup["preset"] = "choice";
			choiceGroup["choices"] = json::array();

			while (i <= rowCount && getCellStr(sheet, i, COL_TYPE) == "choice")
			{
				json choiceItem;
				auto text = getCellStr(sheet, i, COL_TEXT);
				choiceItem["text"] = set_html(text, sheet, i);

				if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "result")
				{
					choiceItem["result"] = json::array();
					i++;

					while (i <= rowCount && getCellStr(sheet, i, COL_TYPE) == "result")
					{
						json line;
						auto speaker = getCellStr(sheet, i, COL_SPEAKER);
						if (speaker.starts_with("image"))
						{
							line["type"] = "image";
							line["src"] = speaker;
							auto caption = getCellStr(sheet, i, COL_TEXT);
							if (!caption.empty())
								line["caption"] = caption;

						}
						else if (speaker.empty())
						{
							line["preset"] = "nar";
							set_line(line, sheet, i);
						}
						else
							set_line(line, sheet, i);

						choiceItem["result"].push_back(line);

						if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "result")
							i++;
						else
							break;
					}
				}
				choiceGroup["choices"].push_back(choiceItem);

				if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "choice")
					i++;
				else 
					break;
			}
			root["lines"].push_back(choiceGroup);
		}
		else if (type.starts_with("image"))
		{
			json line;
			auto caption = getCellStr(sheet, i, COL_TEXT);
			if (!caption.empty())
				line["caption"] = caption;

			line["type"] = "image";
			line["src"] = type;

			root["lines"].push_back(line);
		}
		else if (type == "dialogue")
		{
			json line;
			string speaker = getCellStr(sheet, i, COL_SPEAKER);
			if (speaker.empty())
				line["preset"] = "nar";

			set_line(line, sheet, i);
			root["lines"].push_back(line);
		}
		else if (type == "nar-center")
		{
			json line;
			line["preset"] = "nar-center";
			set_line(line, sheet, i);

			root["lines"].push_back(line);
		}
		else if (type == "contour")
		{
			json line;
			line["type"] = "contour";
			root["lines"].push_back(line);
		}
	}
	return root;
}

void parse(const filesystem::path& file)
{
	Log(file.filename().string());

	OpenXLSX::XLDocument doc(file.string());
	auto wb = doc.workbook();
	for (const auto& sheetName : wb.worksheetNames())
	{
		Log("\t" + sheetName + "...");
		if ('#' == sheetName.front() || sheetName.empty())
		{
			Log("- Skipped.");
			continue;
		}

		auto wks = wb.worksheet(sheetName);
		static constexpr auto minRowCount = 2;
		if (wks.rowCount() < minRowCount)
			throw runtime_error("Needs at least 2 of row counts.");

		auto spreadSheetName = file.filename().replace_extension("").string();
		auto outFile = sheetName + ".json";

		auto path = file;
		path = path.remove_filename();
		path /= spreadSheetName;

		if (!exists(path))
			std::filesystem::create_directory(path);

		ofstream ofs(path / outFile);
		ofs << convert_json(wks);
	}

	doc.close();
}

void parse_all_in_directory(const filesystem::path& path)
{
	for (const auto& entry : filesystem::recursive_directory_iterator(path))
	{
		if (is_regular_file(entry) &&
			entry.path().extension() == ".xlsx" &&
			*entry.path().filename().string().data() != '~')
		{
			parse(entry);
		}
	}
}

int main(int argCount, char* args[])
{
	try
	{
		if (1 == argCount)
		{
			char buffer[1024]{};
			GetCurrentDirectoryA(1024, buffer);
			parse_all_in_directory({ buffer });
		}
		else
		{
			for (auto i = 1; i < argCount; ++i)
				parse(args[i]);
		}
	}
	catch (const exception& e)
	{
		Log(e.what());
	}
}
