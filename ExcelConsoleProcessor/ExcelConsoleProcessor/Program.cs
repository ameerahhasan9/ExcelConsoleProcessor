using ExcelConsoleProcessor;
using ExcelConsoleProcessor.DTO;
using ExcelConsoleProcessor.Model;
using OfficeOpenXml;

namespace ExcelConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome back ^_^");
            bool again = true;
            Console.WriteLine("Please enter the path of the Excel file you want to process..");
            string pathExcel = Console.ReadLine();
            do
            {
                try
                {
                    // connection file Excel
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(pathExcel)))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.Commercial;
                        var sheet = package.Workbook.Worksheets["data"];
                        var users = new Program().GetList<UserDTO>(sheet);
                        if (users != null && users.Count != 0)
                        {
                            // connection database and add records
                            using (var db = new ModelContext())
                            {
                                // Note: This requires the database to be created before running.
                                foreach (var objuser in users)
                                {
                                    // Add a record in the user table                          
                                    db.Add(new User { FullName = objuser.Name, Email = objuser.Email, IsActive = objuser.IsActive });
                                    db.SaveChanges();
                                }
                                // successfully Add a records in the user table
                                Console.WriteLine("\nData inserted successfully ^_^");
                                Console.WriteLine("Total number of users inserted is : " + users.Count);
                            }
                        }
                        else
                        {
                            // Failed to add data in user table
                            Console.WriteLine("try again!\n");
                        }
                    }
                }
                catch (Exception)
                {
                    // Failed to add data in user table because error in path excel
                    Console.WriteLine("\nFailed to add data in table,Please check file path !");
                    Console.WriteLine("try again!");
                }

                // try again or finish 
                Console.WriteLine("\nIf you want to try again please enter the path  ^_^ ");
                Console.WriteLine("If you want to try again later, please enter 1");
                pathExcel = Console.ReadLine();

                // if end program
                if (pathExcel == "1")
                {
                    // end program
                    again = false;
                    Console.WriteLine("\nThank you, I hope you like the service ^_____^");
                }
            } while (again);
        }

        //Get data from excel file
        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            try
            {
                if (sheet != null)
                {
                    if (sheet.Cells[1, 1].Value != null)
                    {
                        var colInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                  new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() });


                        for (int row = 2; row <= sheet.Dimension.Rows; row++)
                        {
                            T obj = (T)Activator.CreateInstance(typeof(T));


                            foreach (var prop in typeof(T).GetProperties())
                            {

                                int col = colInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                                var val = sheet.Cells[row, col].Value;
                                if (val == null)
                                {
                                    Console.WriteLine("\nFailed to add data in table, please check the data is correct!");
                                    return null;
                                }
                                // Check Is Active and convert to boolean
                                else if (prop.Name == "IsActive")
                                {

                                    if (val.ToString() == "Yes")
                                    {
                                        val = 1;
                                    }
                                    else
                                    {
                                        val = 0;
                                    }
                                }
                                var propType = prop.PropertyType;
                                prop.SetValue(obj, Convert.ChangeType(val, propType));
                            }
                            list.Add(obj);
                        }


                    }
                    //excel is empty
                    else
                    {
                        Console.WriteLine("\nThe sheet selected is empty");
                    }
                }
                //error path
                else { Console.WriteLine("\nFailed to add data in table,Please check file path !"); }
            }
            //error data in file excel
            catch (Exception)
            {
                Console.WriteLine("\nFailed to add data in table, please check the data is correct!");
            }

            return list;
        }

    }
}
