namespace GozCommunicator.Managers
{
    public class CellExcel
    {
        public string Column { get; set; }
        public int Row { get; set; }

        public CellExcel(string column, int row)
        {
            Column = column;
            Row = row;
        }
    }
}
