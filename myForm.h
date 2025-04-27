#pragma once

namespace Project4 {
	using namespace System;
	using namespace System::Collections::Generic;
	using namespace System::Data;
	using namespace System::Data::OleDb;
	using namespace System::Windows::Forms;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::Button^ button2;
	private: System::Windows::Forms::Button^ button3;
	protected:

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->button3 = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Location = System::Drawing::Point(39, 28);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(702, 182);
			this->dataGridView1->TabIndex = 0;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(629, 226);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(75, 23);
			this->button1->TabIndex = 1;
			this->button1->Text = L"Добавить";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// button2
			// 
			this->button2->Location = System::Drawing::Point(455, 226);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(104, 23);
			this->button2->TabIndex = 2;
			this->button2->Text = L"Редактировать";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &MyForm::button2_Click);
			// 
			// button3
			// 
			this->button3->Location = System::Drawing::Point(345, 226);
			this->button3->Name = L"button3";
			this->button3->Size = System::Drawing::Size(75, 23);
			this->button3->TabIndex = 3;
			this->button3->Text = L"Читать таблицу";
			this->button3->UseVisualStyleBackColor = true;
			this->button3->Click += gcnew System::EventHandler(this, &MyForm::button3_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(753, 261);
			this->Controls->Add(this->button3);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataGridView1);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}

		// Функция для безопасного преобразования значения в строку, представляющую целое число
		String^ ConvertToSafeIntegerString(Object^ value)
		{
			if (value == nullptr)
			{
				return "0"; // Значение по умолчанию
			}

			String^ stringValue = value->ToString(); // Преобразуем в строку

			int intValue = 0;
			if (Int32::TryParse(stringValue, intValue))
			{
				return intValue.ToString(); // Успешно преобразовано в целое число
			}
			else
			{
				return "0"; // Не удалось преобразовать, возвращаем значение по умолчанию
			}
		}

		String^ GetFirstSheetName(DataTable^ schemaTable)
		{
			if (schemaTable == nullptr || schemaTable->Rows == nullptr)
				return String::Empty;

			// Получаем коллекцию строк
			DataRowCollection^ rows = schemaTable->Rows;

			// Проверяем, есть ли строки
			if (rows->Count == 0)
				return String::Empty;

			// Получаем первую строку
			DataRow^ firstRow = rows[0];
			if (firstRow == nullptr)
				return String::Empty;

			// Получаем значение столбца TABLE_NAME
			Object^ tableNameObj = firstRow["TABLE_NAME"];
			if (tableNameObj == nullptr)
				return String::Empty;

			return tableNameObj->ToString();
		}
void LoadExcelColumnNamesToDataGridView(DataGridView^ dataGridView1, String^ excelFilePath)
		{
			if (dataGridView1 == nullptr || String::IsNullOrEmpty(excelFilePath))
			{
				MessageBox::Show("Некорректные параметры");
				return;
			}
			OleDbConnection^ conn = nullptr;
			try
			{
			// 2. Подключение к Excel
				String^ connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath +	";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";
				conn = gcnew OleDbConnection(connString);
				conn->Open();
				// 3. Получаем список листов
				DataTable^ schemaTable = conn->GetOleDbSchemaTable(OleDbSchemaGuid::Tables, nullptr);
				String^ firstSheetName = GetFirstSheetName(schemaTable);
				if (String::IsNullOrEmpty(firstSheetName))
				{
					MessageBox::Show("Не удалось получить имя листа");
					return;
				}
				// 4. Получаем структуру столбцов
				DataTable^ columnSchema = nullptr;
				OleDbCommand^ cmd = gcnew OleDbCommand("SELECT TOP 1 * FROM [" + firstSheetName + "]", conn);
				OleDbDataReader^ reader = cmd->ExecuteReader();
				columnSchema = reader->GetSchemaTable();
				reader->Close();
				// 5. Обновляем DataGridView
				dataGridView1->Columns->Clear();

				for (int i = 0; i < columnSchema->Rows->Count; i++)
				{
					DataRow^ row = columnSchema->Rows[i];
					String^ columnName = row["ColumnName"]->ToString();
					DataGridViewTextBoxColumn^ column = gcnew DataGridViewTextBoxColumn();
					column->Name = columnName;
					column->HeaderText = columnName;
					dataGridView1->Columns->Add(column);
				}
			}
			catch (Exception^ ex)
			{
				MessageBox::Show("Ошибка: " + ex->Message);
			}
			finally
			{
				if (conn != nullptr && conn->State == ConnectionState::Open)
					conn->Close();
			}
		}
bool SheetExists(OleDbConnection^ conn, String^ sheetName) {
	try {
		OleDbCommand^ cmd = gcnew OleDbCommand(
			"SELECT TOP 1 * FROM [" + sheetName + "]",
			conn
		);
		cmd->ExecuteReader()->Close(); // Попытка прочитать лист
		return true;
	}
	catch (OleDbException^) {
		return false; // Лист не существует
	}
}
#pragma endregion
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
		LoadExcelColumnNamesToDataGridView(dataGridView1, "file.xls");
		DataGridViewTextBoxColumn^ column1 = gcnew DataGridViewTextBoxColumn();
		column1->HeaderText = "ID";
		dataGridView1->Columns->Add(column1);
		dataGridView1->Columns[0]->ReadOnly = true;
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		OleDbConnection^ conn = nullptr;
		try
		{
			// 1. Подключение к Excel (HDR=YES)
			String^ connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=file.xls;Extended Properties=\"Excel 8.0;HDR=YES;\"";
			conn = gcnew OleDbConnection(connString);
			conn->Open();
			String^ sheetName = "Sheet1";

			// Создаем строку для CREATE TABLE
			String^ createTableQuery = String::Format("CREATE TABLE [{0}$] (", sheetName);
			List<String^>^ columnDefinitions = gcnew List<String^>();

			for (int i = 0; i < dataGridView1->Columns->Count; i++)
			{
				DataGridViewColumn^ column = dataGridView1->Columns[i];
				String^ columnName = column->HeaderText;
				// Экранируем имя столбца, если оно содержит пробелы или специальные символы
				if (columnName->Contains(" ") || columnName->Contains("-") || columnName->Contains("."))
				{
					columnName = "[" + columnName + "]";
				}

				String^ columnType = "VARCHAR(255)"; // Тип данных по умолчанию
				if (column->ValueType == int::typeid || column->HeaderText->Contains("Возраст") || column->HeaderText->Contains("ID"))
				{
					columnType = "INT";
				}
				else if (column->ValueType == DateTime::typeid)
				{
					columnType = "DATETIME";
				}
				columnDefinitions->Add(columnName + " " + columnType);
			}
			createTableQuery += String::Join(",", columnDefinitions) + ")";
			OleDbCommand^ createTableCmd = gcnew OleDbCommand(createTableQuery, conn);
			createTableCmd->ExecuteNonQuery();

			// 3. Вставляем данные (без вставки заголовков)
			for (int i = 0; i < dataGridView1->Rows->Count; i++)
			{
				DataGridViewRow^ row = dataGridView1->Rows[i];

				if (!row->IsNewRow)
				{
					String^ insertQuery = String::Format("INSERT INTO [{0}$] VALUES (", sheetName);
					List<String^>^ parameterNames = gcnew List<String^>();
					for (int j = 0; j < row->Cells->Count; j++)
					{
						parameterNames->Add(String::Format("@p{0}", j));
						if (j > 0) insertQuery += ",";
						insertQuery += String::Format("@p{0}", j);
					}
					insertQuery += ")";

					OleDbCommand^ insertCmd = gcnew OleDbCommand(insertQuery, conn);

					// Добавляем параметры
					for (int j = 0; j < row->Cells->Count; j++)
					{
						DataGridViewCell^ cell = row->Cells[j];
						DataGridViewColumn^ column = dataGridView1->Columns[j];
						String^ value;

						if (column->ValueType == int::typeid || column->HeaderText->Contains("id") || column->HeaderText->Contains("Возраст"))
						{
							value = ConvertToSafeIntegerString(cell->Value);
							insertCmd->Parameters->AddWithValue(String::Format("@p{0}", j), Convert::ToInt32(value));
						}
						else
						{
							value = cell->Value == nullptr ? "" : cell->Value->ToString();
							// Экранируем апострофы
							value = value->Replace("'", "''");
							insertCmd->Parameters->AddWithValue(String::Format("@p{0}", j), value);
						}
					}
					insertCmd->ExecuteNonQuery();
				}
			}

			MessageBox::Show("Данные успешно экспортированы!");
		}
		catch (Exception^ ex)
		{
			MessageBox::Show("Ошибка: " + ex->Message);
		}
		finally
		{
			if (conn != nullptr && conn->State == ConnectionState::Open)
				conn->Close();
		}
	}

private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {

	OleDbConnection^ conn = nullptr;
    try
    {
        // 1. Подключение к Excel
        String^ connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=file.xls;Extended Properties=\"Excel 8.0;HDR=YES;\"";
        conn = gcnew OleDbConnection(connString);
        conn->Open();
        String^ sheetName = "Sheet1";
        String^ keyColumnName = "ID"; // Укажите имя ключевого столбца

        // Цикл по строкам DataGridView
        for (int i = 0; i < dataGridView1->Rows->Count; i++)
        {
            DataGridViewRow^ row = dataGridView1->Rows[i];


            if (!row->IsNewRow)
            {
                // Получаем значение ключевого столбца
				int keyValue;
				Object^ cellValue = row->Cells["ID"]->Value; // Замените "ID" на имя вашего ключевого столбца
				if (cellValue == nullptr)
				{
					keyValue = 0; // Значение по умолчанию
				}
				else
				{
					String^ stringValue = cellValue->ToString();
					if (!Int32::TryParse(stringValue, keyValue))
					{
						// Обработка ошибки преобразования
						MessageBox::Show("Не удалось преобразовать в int: " + stringValue);
						keyValue = 0; // Значение по умолчанию
					}
				}
                // 1. Пытаемся обновить существующую строку
                String^ updateQuery = String::Format("UPDATE [{0}$] SET ", sheetName);
                List<String^>^ setClauses = gcnew List<String^>();

                for (int j = 0; j < dataGridView1->Columns->Count; j++)
                {
                    DataGridViewColumn^ column = dataGridView1->Columns[j];
                    String^ columnName = column->HeaderText;

                    if (columnName != keyColumnName)
                    {
                        String^ value;
                        String^ parameterName = String::Format("@p{0}", j);

                        if (column->ValueType == int::typeid || columnName == "Возраст" || columnName == "ID") // Проверка на "Возраст"
                        {
                            value = ConvertToSafeIntegerString(row->Cells[j]->Value);
                            setClauses->Add(String::Format("[{0}] = {1}", columnName, parameterName)); // Без кавычек
                        }
                        else
                        {
                            value = row->Cells[j]->Value == nullptr ? "" : row->Cells[j]->Value->ToString();
                            value = value->Replace("'", "''"); // Экранируем апострофы
                            setClauses->Add(String::Format("[{0}] = {1}", columnName, parameterName)); // С кавычками

                        }
                    }
                }
                updateQuery += String::Join(",", setClauses);
				// Добавляем условие WHERE
                updateQuery += String::Format(" WHERE [{0}] = @key", keyColumnName);

                // Создаем OleDbCommand
                OleDbCommand^ updateCmd = gcnew OleDbCommand(updateQuery, conn);

                // Добавляем параметры
                for (int j = 0; j < dataGridView1->Columns->Count; j++)
                {
                    DataGridViewColumn^ column = dataGridView1->Columns[j];
                    String^ columnName = column->HeaderText;

                    if (columnName != keyColumnName)
                    {
                        String^ value;
                        String^ parameterName = String::Format("@p{0}", j);

                        if (column->ValueType == int::typeid || columnName == "Возраст" || columnName == "ID") // Проверка на "Возраст"
                        {
                            value = ConvertToSafeIntegerString(row->Cells[j]->Value);
                            updateCmd->Parameters->AddWithValue(parameterName, Convert::ToInt32(value)); // Числовой параметр
                        }
                        else
                        {
                            value = row->Cells[j]->Value == nullptr ? "" : row->Cells[j]->Value->ToString();
                            // Экранируем апострофы
                            value = value->Replace("'", "''");
                            updateCmd->Parameters->AddWithValue(parameterName, value); // Строковый параметр
                        }
                    }
                }

                updateCmd->Parameters->AddWithValue("@key", keyValue); // Добавляем параметр для ключевого столбца

                // Выполняем UPDATE запрос
                int rowsAffected = updateCmd->ExecuteNonQuery();

                // Если ни одна строка не была обновлена, выводим сообщение
                if (rowsAffected == 0)
                {
                    MessageBox::Show(String::Format("Значения столбца {0} = {1} нельзя изменять. Строки не были обновлены ", keyColumnName, keyValue));
                }
            }
        }

        MessageBox::Show("ОК!");
    }
    catch (Exception^ ex)
    {
        MessageBox::Show("Ошибка: " + ex->Message);
    }
    finally
    {
        if (conn != nullptr && conn->State == ConnectionState::Open)
            conn->Close();
    }
	
}
private: System::Void button3_Click(System::Object^ sender, System::EventArgs^ e) {
	String^ sheetName = "Sheet1";
		OleDbConnection^ conn = nullptr;
		try
		{
			String^ connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= file.xls ;Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";
			conn = gcnew OleDbConnection(connString);
			conn->Open();
			DataTable^ schemaTable = conn->GetOleDbSchemaTable(OleDbSchemaGuid::Tables, nullptr);
			bool sheetExists = false;
			for each (DataRow ^ row in schemaTable->Rows)
			{
				if (row["TABLE_NAME"]->ToString()->Equals(sheetName + "$"))
				{
					sheetExists = true;
					break;
				}
			}
			if (!sheetExists)
			{
				MessageBox::Show("Лист " + sheetName + " не найден в файле Excel!");
				return;
			}

			dataGridView1->Columns->Clear();
			dataGridView1->Rows->Clear();
			dataGridView1->Refresh();

			String^ selectQuery = "SELECT * FROM [" + sheetName + "$]";
			OleDbCommand^ selectCmd = gcnew OleDbCommand(selectQuery, conn);
			OleDbDataAdapter^ adapter = gcnew OleDbDataAdapter(selectCmd);
			DataTable^ dataTable = gcnew DataTable();

			adapter->Fill(dataTable);
			dataGridView1->DataSource = dataTable;
			dataGridView1->AutoResizeColumns();

			MessageBox::Show("Данные успешно загружены из Excel!");
		}
		catch (Exception^ ex)
		{
			MessageBox::Show("Ошибка при загрузке из Excel: " + ex->Message);
		}
		finally
		   {
			   if (conn != nullptr && conn->State == ConnectionState::Open)
			   {
				   conn->Close();
			   }
		   }
}
};
}
