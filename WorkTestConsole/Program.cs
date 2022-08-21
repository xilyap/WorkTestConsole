using ExcelDataReader;
using Microsoft.Data.Sqlite;
using System.Text;
using WorkTestConsole;
using Xceed.Words.NET;

internal class Program
{
    const string CREATE_TABLE_QUERY = "CREATE TABLE IF NOT EXISTS \"users\" " +
                "(\"id\"INTEGER NOT NULL," +
                "\"family\"TEXT NOT NULL," +
                "\"name\"\tTEXT NOT NULL," +
                "\"gender\"\tTEXT NOT NULL," +
                "\"age\"\tINTEGER NOT NULL," +
                "\"status\"\tTEXT NOT NULL," +
                "PRIMARY KEY(\"id\" AUTOINCREMENT))";
    const string COUNT_MALE_FEMALE_QUERY = "SELECT count(id) from users where gender = \"м\" or gender = \"ж\"";
    const string COUNT_MALE_MIDDLE_AGE_QUERY = "SELECT count(id) FROM users WHERE gender = \"м\" AND age BETWEEN 30 AND 40";
    const string COUNT_STANDART_PREMIUM_ACCOUNTS_QUERY = "SELECT count(id) from users where status = \"стандарт\" or status = \"премиум\"";
    const string COUNT_FEMALE_PREMIUM_BEFORE_MIDDLE_AGE_QUERY = "SELECT count(id) FROM users WHERE gender = \"ж\" AND age < 30";
    const string INSERT_USER_QUERY = "INSERT INTO users (family, name, gender, age, status)" +
                    " VALUES (\"{0}\", \"{1}\",\"{2}\", {3}, \"{4}\"); ";

    private static void Main(string[] args)
    {

        IniFile ini = new IniFile("config.ini");
        if (ini.KeyExists("FilePath"))
        {
            string outputPath;
            if (ini.KeyExists("OutputFolder"))
            {
                outputPath = ini.Read("OutputFolder") + "/out.docx";
            }
            else
            {
                ini.Write("OutputFolder", "");
                outputPath = "out.docx";
            }
            var filePath = ini.Read("FilePath");
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                ProcessFile(filePath);
                GenerateReport(outputPath);
                Console.WriteLine("Готово!");
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        else
        {
            Console.WriteLine("Файл config.ini испорчен, или содержит не все объекты.\n\r" +
                "Поле FilePath является обязательным!\n\r" +
                "Заполните недостающие поля и запустите программу заново!");
            if (!ini.KeyExists("FilePath")) ini.Write("FilePath", "");
        }
        Console.WriteLine("Нажмите любую клавишу, чтобы закрыть программу");
        Console.ReadKey();
    }

    /// <summary>
    /// Метод, генерирующий отчёт по базе данных SQLite, предварительно созданной в методе <see cref="ProcessFile"/>
    /// </summary>
    /// <param name="outputPath">Путь в директорию вывода</param>
    private static void GenerateReport(string outputPath)
    {
        long[] result = new long[4];

        using (var connection = new SqliteConnection("Data Source = database.db"))
        {
            connection.Open();
            SqliteCommand command = new SqliteCommand(COUNT_MALE_FEMALE_QUERY, connection);
            result[0] = (long)command.ExecuteScalar();
            command = new SqliteCommand(COUNT_MALE_MIDDLE_AGE_QUERY, connection);
            result[1] = (long)command.ExecuteScalar();
            command = new SqliteCommand(COUNT_STANDART_PREMIUM_ACCOUNTS_QUERY, connection);
            result[2] = (long)command.ExecuteScalar();
            command = new SqliteCommand(COUNT_FEMALE_PREMIUM_BEFORE_MIDDLE_AGE_QUERY, connection);
            result[3] = (long)command.ExecuteScalar();
        }
        DocX doc = DocX.Create(outputPath);
        doc.InsertParagraph("Отчёт по документу:").Bold();
        doc.InsertParagraph(String.Format("1. мужчин и женщин = {0}\r\n" +
            "2. мужчин в возрасте 30-40 лет = {1}\r\n" +
            "3. стандартных и премиум-аккаунтов = {2}\r\n" +
            "4. женщин с премиум-аккаунтом в возрасте до 30 лет = {3}", result[0], result[1], result[2], result[3]));
        doc.Save();
    }

    /// <summary>
    /// Метод, предназначенный для обработки .xlsx файла, заданного согласно заданию. В результе создаёт локальную базу данных SQLite с содержимым обрабатываемого файла.
    /// </summary>
    /// <param name="filePath">Путь к входному файлу</param>
    private static void ProcessFile(string filePath)
    {
        FileStream fstrem = File.Open(filePath, FileMode.Open);
        ExcelReaderConfiguration conf = new ExcelReaderConfiguration();
        conf.FallbackEncoding = Encoding.GetEncoding(1252);
        IExcelDataReader reader = ExcelReaderFactory.CreateReader(fstrem, conf);
        File.Delete("database.db");
        using (var connection = new SqliteConnection("Data Source = database.db"))
        {
            connection.Open();
            using (var transaction = connection.BeginTransaction())
            {
                SqliteCommand command = connection.CreateCommand();
                command.CommandText = CREATE_TABLE_QUERY;
                command.Transaction = transaction;
                command.ExecuteNonQuery();
                reader.Read();
                for (int i = 0; i < reader.RowCount; i++)
                {
                    reader.Read();
                    string query = String.Format(INSERT_USER_QUERY, reader[0], reader[1], reader[2], reader[3], reader[4]);
                    using (var cmd = new SqliteCommand(query, connection, transaction))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                transaction.Commit();
            }
        }
    }
}