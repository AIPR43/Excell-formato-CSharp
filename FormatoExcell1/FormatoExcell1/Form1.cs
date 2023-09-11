using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace FormatoExcell1
{

    public partial class Form1 : Form
    {
        private System.Data.DataTable dataTable;
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = "Data Source=(localdb)\\AIPR;Initial Catalog=MBDD1;Integrated Security=true;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                // Ejecuta tu consulta SQL y llena el DataTable.
                string query = "SELECT * FROM Clientes"; // Utiliza el nombre de tu tabla "Clientes".
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                dataTable = new System.Data.DataTable();
                adapter.Fill(dataTable);
            }

            // Crear una instancia de Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

            // Supongamos que tienes un DataTable llamado "dataTable" con tus datos.
            int fila = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                worksheet.Cells[fila, 1] = row["IDCliente"];
                worksheet.Cells[fila, 2] = row["Nombre"];
                worksheet.Cells[fila, 3] = row["Apellido"];
                worksheet.Cells[fila, 4] = row["Monto"];
                // Continúa colocando los datos en las celdas correspondientes.
                fila++;
            }

            // Cálculo de sumatorias
            worksheet.Cells[fila, 1] = "Total:";
            worksheet.Cells[fila, 4].Formula = "SUM(D1:D" + (fila - 1) + ")";

            // Mostrar números de línea
            int contador = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                worksheet.Cells[fila, 1] = contador;
                worksheet.Cells[fila, 2] = row["Nombre"];
                worksheet.Cells[fila, 3] = row["Apellido"];
                // Continúa colocando los datos en las celdas correspondientes.
                fila++;
                contador++;
            }

            // Guardar el archivo Excel
            workbook.SaveAs("C:\\Users\\venta\\source\\repos\\FormatoExcell1\\ruta_del_archivo.xlsx");

            // Abrir el archivo en modo solo lectura
            workbook.ReadOnlyRecommended = true;
            workbook.Close(false);
            excelApp.Quit();
            dataGridView1.DataSource = dataTable;
        }
    }
}
