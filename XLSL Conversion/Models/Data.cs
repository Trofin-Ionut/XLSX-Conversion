namespace XLSL_Conversion.Models
{
    public class Data
    {
        public Dictionary<string, int> sum1 = new();
        public Dictionary<string, int> sum2 = new();
        public Dictionary<string, int> sum3 = new();
        public void AddStorage(string storage)
        {
            sum1.TryAdd(storage, 0);
            sum2.TryAdd(storage, 0);
            sum3.TryAdd(storage, 0);
        }
    }
}
