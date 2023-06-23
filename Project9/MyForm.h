#pragma once
namespace Project9 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Microsoft::Office::Interop::Excel;
	using namespace System::IO;

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

	protected:
	private: System::Windows::Forms::OpenFileDialog^ openFileDialog1;
	private: System::Windows::Forms::Button^ loadButton;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::ComboBox^ comboBoxCulture;














	private: System::Windows::Forms::ComboBox^ comboBoxAreaNumber;

	private: System::Windows::Forms::Label^ label4;
	private: System::Windows::Forms::Label^ label5;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ ColumnWeek;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ ColumnValueNDVI;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ ColumnMinNDVI;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ ColumnMark;











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
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle1 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->openFileDialog1 = (gcnew System::Windows::Forms::OpenFileDialog());
			this->loadButton = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->comboBoxCulture = (gcnew System::Windows::Forms::ComboBox());
			this->comboBoxAreaNumber = (gcnew System::Windows::Forms::ComboBox());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->label5 = (gcnew System::Windows::Forms::Label());
			this->ColumnWeek = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->ColumnValueNDVI = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->ColumnMinNDVI = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->ColumnMark = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// openFileDialog1
			// 
			this->openFileDialog1->FileName = L"openFileDialog1";
			// 
			// loadButton
			// 
			this->loadButton->BackColor = System::Drawing::Color::DarkSeaGreen;
			this->loadButton->Font = (gcnew System::Drawing::Font(L"Palatino Linotype", 10.2F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->loadButton->Location = System::Drawing::Point(12, 12);
			this->loadButton->Name = L"loadButton";
			this->loadButton->Size = System::Drawing::Size(179, 50);
			this->loadButton->TabIndex = 1;
			this->loadButton->Text = L"загрузить .csv";
			this->loadButton->UseVisualStyleBackColor = false;
			this->loadButton->Click += gcnew System::EventHandler(this, &MyForm::loadButton_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->BackgroundColor = System::Drawing::SystemColors::GradientInactiveCaption;
			dataGridViewCellStyle1->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle1->BackColor = System::Drawing::SystemColors::GradientInactiveCaption;
			dataGridViewCellStyle1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 7.8F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			dataGridViewCellStyle1->ForeColor = System::Drawing::SystemColors::WindowText;
			dataGridViewCellStyle1->SelectionBackColor = System::Drawing::SystemColors::Info;
			dataGridViewCellStyle1->SelectionForeColor = System::Drawing::SystemColors::InfoText;
			dataGridViewCellStyle1->WrapMode = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView1->ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(4) {
				this->ColumnWeek,
					this->ColumnValueNDVI, this->ColumnMinNDVI, this->ColumnMark
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 68);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->RowHeadersWidth = 51;
			this->dataGridView1->RowTemplate->Height = 24;
			this->dataGridView1->Size = System::Drawing::Size(1273, 497);
			this->dataGridView1->TabIndex = 2;
			// 
			// comboBoxCulture
			// 
			this->comboBoxCulture->BackColor = System::Drawing::SystemColors::GradientInactiveCaption;
			this->comboBoxCulture->FormattingEnabled = true;
			this->comboBoxCulture->Items->AddRange(gcnew cli::array< System::Object^  >(11) {
				L"Озимая пшеница\t", L"Озимая рожь", L"Нут",
					L"Соя на орошении, богаре", L"Соя на богаре", L"Ячмень", L"Яровая пшеница", L"Суданская трава", L"Просо", L"Сорго", L"Подсолнечник"
			});
			this->comboBoxCulture->Location = System::Drawing::Point(197, 38);
			this->comboBoxCulture->Name = L"comboBoxCulture";
			this->comboBoxCulture->Size = System::Drawing::Size(272, 24);
			this->comboBoxCulture->TabIndex = 9;
			this->comboBoxCulture->SelectedIndexChanged += gcnew System::EventHandler(this, &MyForm::comboBoxCulture_SelectedIndexChanged);
			// 
			// comboBoxAreaNumber
			// 
			this->comboBoxAreaNumber->BackColor = System::Drawing::SystemColors::GradientInactiveCaption;
			this->comboBoxAreaNumber->FormattingEnabled = true;
			this->comboBoxAreaNumber->Location = System::Drawing::Point(475, 38);
			this->comboBoxAreaNumber->Name = L"comboBoxAreaNumber";
			this->comboBoxAreaNumber->Size = System::Drawing::Size(811, 24);
			this->comboBoxAreaNumber->TabIndex = 10;
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Font = (gcnew System::Drawing::Font(L"Palatino Linotype", 12, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label4->Location = System::Drawing::Point(244, 8);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(172, 27);
			this->label4->TabIndex = 11;
			this->label4->Text = L"Список культур";
			// 
			// label5
			// 
			this->label5->AutoSize = true;
			this->label5->Font = (gcnew System::Drawing::Font(L"Palatino Linotype", 12, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label5->Location = System::Drawing::Point(872, 8);
			this->label5->Name = L"label5";
			this->label5->Size = System::Drawing::Size(89, 27);
			this->label5->TabIndex = 12;
			this->label5->Text = L"№ поля";
			// 
			// ColumnWeek
			// 
			this->ColumnWeek->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::Fill;
			this->ColumnWeek->HeaderText = L"Неделя";
			this->ColumnWeek->MinimumWidth = 6;
			this->ColumnWeek->Name = L"ColumnWeek";
			// 
			// ColumnValueNDVI
			// 
			this->ColumnValueNDVI->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::Fill;
			this->ColumnValueNDVI->HeaderText = L"Показатель NDVI";
			this->ColumnValueNDVI->MinimumWidth = 6;
			this->ColumnValueNDVI->Name = L"ColumnValueNDVI";
			// 
			// ColumnMinNDVI
			// 
			this->ColumnMinNDVI->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::Fill;
			this->ColumnMinNDVI->HeaderText = L"min NDVI";
			this->ColumnMinNDVI->MinimumWidth = 6;
			this->ColumnMinNDVI->Name = L"ColumnMinNDVI";
			// 
			// ColumnMark
			// 
			this->ColumnMark->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::Fill;
			this->ColumnMark->HeaderText = L"Оценка";
			this->ColumnMark->MinimumWidth = 6;
			this->ColumnMark->Name = L"ColumnMark";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::LemonChiffon;
			this->BackgroundImage = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"$this.BackgroundImage")));
			this->ClientSize = System::Drawing::Size(1297, 591);
			this->Controls->Add(this->label5);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->comboBoxAreaNumber);
			this->Controls->Add(this->comboBoxCulture);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->loadButton);
			this->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 7.8F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"MyForm";
			this->RightToLeft = System::Windows::Forms::RightToLeft::No;
			this->Text = L"NDVI";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}

#pragma endregion

		int countTable = 0;

	private: System::Void loadButton_Click(System::Object^ sender, System::EventArgs^ e) {
		openFileDialog1->Filter = "(excel file(*.csv)| *.csv";
		openFileDialog1->RestoreDirectory = true;
		System::Data::DataTable^ tb = gcnew System::Data::DataTable();
		try
		{
			if (openFileDialog1->ShowDialog() == System::Windows::Forms::DialogResult::OK)
			{
				array<String^>^ rows = System::IO::File::ReadAllLines(openFileDialog1->FileName, System::Text::Encoding::Default);

				//нужно поправить добавление строк (неверно)
				if (dataGridView1->RowCount < 1) {
					dataGridView1->Rows->Add(rows->Length - 3);
				}
				else if (dataGridView1->RowCount < (rows->Length - 3)) {
					dataGridView1->Rows->Add(rows->Length - 3 - dataGridView1->RowCount);
				}
				else if (dataGridView1->RowCount > (rows->Length - 3)) {
					int count = dataGridView1->RowCount - (rows->Length - 3);
					for (int i = 0; i < count; i++) {
						dataGridView1->Rows->Remove(dataGridView1->Rows[0]);
					}
				}

				String^ area = rows[0]->Split(';')[1]->ToString();
				comboBoxAreaNumber->Items->Add(area);
				comboBoxAreaNumber->SelectedIndex = countTable;

				array<String^>^ cells;
				for (int j = 3; j < rows->Length; j++)
				{
					cells = rows[j]->Split(';');
					for (int i = 0; i < cells->Length; i++)
					{
						if (cells[i]->Contains(".86")) {
							cells[i] = cells[i]->Remove(cells[i]->IndexOf('.'));
						}
						dataGridView1->Rows[j - 3]->Cells[i]->Value = cells[i];
					}
				}
			}
			countTable++;
		}
		catch (Exception^ ex) {
			MessageBox::Show(this, "не удалось открыть файл", "ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}

		   array<double^, 2>^ minNDVI;
		   int week = 52;
		   int countCulture = 10;
private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
	minNDVI = gcnew array<double^, 2>(countCulture, week) {
		{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.3, 0.3, 0.3, 0.3, 0.4, 0.5, 0.6, 0.6, 0.7, 0.7, 0.8, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0},
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.4, 0.5, 0.5, 0.5, 0.6, 0.6, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.7, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.4, 0.5, 0.5, 0.6, 0.5, 0.4, 0.4, 0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.3, 0.4, 0.4, 0.5, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.5, 0.4, 0.4, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.4, 0.4, 0.5, 0.5, 0.5, 0.5, 0.5, 0.4, 0.4, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.4, 0.4, 0.5, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.5, 0.4, 0.4, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.4, 0.5, 0.6, 0.5, 0.5, 0.5, 0.4, 0.4, 0.4, 0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.4, 0.5, 0.5, 0.6, 0.6, 0.5, 0.5, 0.5, 0.5, 0.5, 0.51, 0.5, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.4, 0.5, 0.5, 0.6, 0.6, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.4, 0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0 },
		{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.3, 0.3, 0.4, 0.5, 0.6, 0.7, 0.6, 0.5, 0.4, 0.4, 0.4, 0.4, 0.3, 0.3, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.0, 0.0, 0.0, 0.0, 0.0 },
	};
}

private: System::Void comboBoxCulture_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {

	int index = comboBoxCulture->SelectedIndex;
	int startWeek = 0;
	int maxWeek = 52;
	DataGridViewCellStyle^ rowRed = gcnew DataGridViewCellStyle();
	rowRed->BackColor = Color::Red;
	DataGridViewCellStyle^ rowGreen = gcnew DataGridViewCellStyle();
	rowGreen->BackColor = Color::Green;

	if (dataGridView1->RowCount < 1)
	{
		dataGridView1->Rows->Add(maxWeek);
	}
	else if (dataGridView1->RowCount > 1) {
		startWeek = Convert::ToInt32(dataGridView1->Rows[0]->Cells[0]->Value);
		maxWeek = startWeek + dataGridView1->RowCount;
	}

	for (int i = startWeek; i < maxWeek; i++)
	{
		dataGridView1->Rows[i - startWeek]->Cells[2]->Value = minNDVI[index, i];
		double mark = Convert::ToDouble(dataGridView1->Rows[i - startWeek]->Cells[1]->Value) - Convert::ToDouble(minNDVI[index, i]);
		dataGridView1->Rows[i - startWeek]->Cells[3]->Value = mark;
		if (mark < 0.0) {
			dataGridView1->Rows[i - startWeek]->DefaultCellStyle = rowRed;
		}
		else {
			dataGridView1->Rows[i - startWeek]->DefaultCellStyle = rowGreen;
		}
		}
}
};
}