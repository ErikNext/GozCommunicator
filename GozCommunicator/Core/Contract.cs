namespace GozCommunicator.Core
{
    internal class Contract
    {
        public string Id { get; set; }

        public string Customer { get; set; }

        public string Theme { get; set; }

        public string NumberGosContract { get; set; }

        private string igk;

        public string Igk
        {
            get
            {
                return igk;
            }

            set
            {
                if (value == "")
                {
                    igk = NumberGosContract;
                }
                else
                {
                    var igkNumber = value.Substring(value.Length - 1);
                    igk = $"ИГК № {igkNumber}";
                }
            }
        }

        public string CustomersСurrentAccountNumber { get; set; }

        private string accountNumberAvionika;

        public string AccountNumberAvionika
        {
            get
            {
                return accountNumberAvionika;
            }
            set
            {
                accountNumberAvionika = value.Substring(0, 20);
            }
        }
        public string Remark { get; set; }
    }
}
