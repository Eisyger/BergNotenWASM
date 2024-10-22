using System.Reflection;

namespace BergNotenWASM.Interfaces
{
    public interface IExportable
    {
        public abstract List<PropertyInfo> GetProperties();
    }
}