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

json create_line_object(const ScriptLine& line)
{
	json obj;

	if (line.type == "image")
	{
		obj["type"] = "image";
		obj["src"] = line.text;
	}
	else
	{
		// 나레이션 판정 (speaker가 없으면 nar)
		if (line.speaker.empty())
		{
			obj["preset"] = "nar";
		}
		else
		{
			obj["speaker"] = line.speaker;
		}

		obj["text"] = line.text;
		if (line.italic) obj["italic"] = true;
		if (line.bold) obj["bold"] = true;
		if (line.size > 0) obj["size"] = line.size;
	}
	return obj;
}

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
regex endColor("</red>|</org>|</blue>");

string change_color(const string& str)
{
	auto text = str;
	text = std::regex_replace(text, red, string("[color=#ff2222]"));
	text = std::regex_replace(text, orange, string("[color=#E48F5D]"));
	text = std::regex_replace(text, blue, string("[color=#507de6]"));
	text = std::regex_replace(text, endColor, string("[/color]"));
	return text;
}


const int COL_TYPE = 1;
const int COL_SPEAKER = 2;
const int COL_TEXT = 3;
const int COL_ITALIC = 4;
const int COL_BOLD = 5;
const int COL_SIZE = 6;
void set_text_and_option(json& j, OpenXLSX::XLWorksheet& sheet, int row)
{
	j["text"] = change_color(getCellStr(sheet, row, COL_TEXT));

	if (getCellInt(sheet, row, COL_ITALIC) == 1)
		j["italic"] = true;
	if (getCellInt(sheet, row, COL_BOLD) == 1)
		j["bold"] = true;
}

json convert_json(OpenXLSX::XLWorksheet& sheet)
{
	json root;
	root["lines"] = json::array();

	int rowCount = sheet.rowCount();

	auto set_text = [&](json& line, int i)
	{
		set_text_and_option(line, sheet, i);
		int size = getCellInt(sheet, i, COL_SIZE);
		if (size > 0)
			line["size"] = size;

		root["lines"].push_back(line);
	};


	for (int i = 2; i <= rowCount; ++i)
	{
		std::string type = getCellStr(sheet, i, COL_TYPE);
		if (type.empty())
			type = "dialogue";

		if (type == "choice")
		{
			json choiceGroup;
			choiceGroup["preset"] = "choice";
			choiceGroup["choices"] = json::array();

			// 연속된 선택지들 처리
			while (i <= rowCount && getCellStr(sheet, i, COL_TYPE) == "choice")
			{
				json choiceItem;
				choiceItem["text"] = getCellStr(sheet, i, COL_TEXT);

				// choice 다음 행에 result가 있는지 확인
				if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "result")
				{
					choiceItem["result"] = json::array();
					i++;

					// result에서도 speaker: text 형식으로 처리
					while (i <= rowCount && getCellStr(sheet, i, COL_TYPE) == "result")
					{
						json resLine;
						std::string spk = getCellStr(sheet, i, COL_SPEAKER);

						if (spk.empty())
							resLine["preset"] = "nar";
						else
							resLine["speaker"] = spk;

						
						set_text_and_option(resLine, sheet, i);
						choiceItem["result"].push_back(resLine);

						// 다음 행도 result면 이어서 진행
						if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "result")
							i++;
						else
							break;
					}
				}
				choiceGroup["choices"].push_back(choiceItem);

				// 다음 행이 또 choice인지 확인
				if (i + 1 <= rowCount && getCellStr(sheet, i + 1, COL_TYPE) == "choice")
					i++;
				else 
					break;
			}
			root["lines"].push_back(choiceGroup);
		}
		else if (type.starts_with("image"))
		{
			auto caption = getCellStr(sheet, i, COL_TEXT);
			root["lines"].push_back({
				{"type", "image"},
				{"src", type},
				{"caption", caption}
				});
		}
		else if (type == "dialogue")
		{
			json line;
			std::string spk = getCellStr(sheet, i, COL_SPEAKER);

			if (spk.empty())
				line["preset"] = "nar";
			else
				line["speaker"] = spk;

			set_text(line, i);
		}
		else if (type == "nar-center")
		{
			json line;
			line["preset"] = "nar-center";
			set_text(line, i);
		}
		else if (type == "contour")
		{
			json line;
			line["type"] = "contour";
			set_text(line, i);
		}
	}
	return root;
}

void parse(const filesystem::path& file)
{
	Log(file.filename().string());

	OpenXLSX::XLDocument doc(file.filename().string());
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
		auto outFile = format("{}_{}.json", spreadSheetName, sheetName);

		ofstream ofs(outFile);
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
			parse(entry.path().filename());
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
