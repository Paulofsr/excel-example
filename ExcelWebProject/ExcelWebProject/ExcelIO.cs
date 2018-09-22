using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWebProject
{
    public class ExcelIO<T>
    {
        private Dictionary<string, MyPropertyInfo> Propriedades { get; set; }
        private Dictionary<string, MyPropertyInfo> PropriedadesComDefault { get; set; }

        public ExcelIO()
        {
            MontarLista();
        }

        private void MontarLista()
        {
            Propriedades = new Dictionary<string, MyPropertyInfo>();
            PropriedadesComDefault = new Dictionary<string, MyPropertyInfo>();
            foreach (PropertyInfo property in typeof(T).GetProperties())
            {
                var dd = property.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;
                if (dd != null)
                {
                    var value = property.GetCustomAttribute(typeof(DefaultValue)) as DefaultValue;
                    var prop = new MyPropertyInfo()
                    {
                        Property = property,
                        DefaultValue = value != null ? value.Value : null
                    };
                    Propriedades.Add(dd.Name.ToUpper(), prop);
                    if(prop.DefaultValue != null)
                        PropriedadesComDefault.Add(dd.Name.ToUpper(), prop);
                }
            }
        }

        public List<T> LerArquivo(string arquivo, string planilha, out List<ExcelError> listaDeErros)
        {
                List<T> lista = new List<T>();
                listaDeErros = new List<ExcelError>();
            try
            {

                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(arquivo, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheet(planilha);

                for (int i = 1; i <= 10005; i++)
                {
                    if (sheet.GetRow(i) != null)
                    {
                        var linha = sheet.GetRow(i);
                        var objeto = MontarObjetoDefault();
                        for (int j = 0; j < linha.LastCellNum; j++)
                        {
                            if (linha.GetCell(j) != null)
                            {
                                try
                                {
                                    var nomeColuna = sheet.GetRow(0).GetCell(j).StringCellValue;
                                    var celula = linha.GetCell(j);
                                    SetValue(objeto, nomeColuna, celula);
                                }
                                catch (Exception ex)
                                {
                                    listaDeErros.Add(new ExcelError()
                                    {
                                        linha = i,
                                        coluna = j,
                                        excecao = ex
                                    });
                                }
                            }
                        }
                        lista.Add(objeto);
                    }
                }
            }
            catch { }
            //finally {File.Delete(arquivo); }
                return lista;
        }

        private T MontarObjetoDefault()
        {
            var result = (T)Activator.CreateInstance(typeof(T));
            foreach (var props in PropriedadesComDefault)
            {
                var propriedade = PropriedadesComDefault[props.Key.ToUpper()];
                if (propriedade != null)
                    propriedade.Property.SetValue(result, GetValor(propriedade, propriedade.DefaultValue));
            }
            return result;
        }

        private void SetValue(object objeto, string displayName, ICell celula)
        {
            var propriedade = Propriedades[displayName.ToUpper()];
            if (propriedade != null)
                propriedade.Property.SetValue(objeto, GetValor(propriedade, celula.ToString()));
        }

        private object GetValor(MyPropertyInfo propriedade, object valor)
        {
            switch (propriedade.Property.PropertyType.FullName)
            {
                case "System.String":
                    return valor.ToString();
                case "System.Decimal":
                    return Convert.ToDecimal(valor.ToString());
                case "System.Guid":
                    return new Guid(valor.ToString());
            }
            return null;
        }

        //private MyPropertyInfo GetProperty(string displayName)
        //{
        //    foreach (PropertyInfo property in typeof(T).GetProperties())
        //    {
        //        //MemberInfo property = typeof(Produto).GetProperties().GetProperty("CliendId");
        //        var dd = property.GetCustomAttribute(typeof(DisplayAttribute)) as DisplayAttribute;
        //        if (dd != null)
        //        {
        //            if (dd.Name.ToUpper() == displayName.ToUpper())
        //            {
        //                var value = property.GetCustomAttribute(typeof(DefaultValue)) as DefaultValue;
        //                return new MyPropertyInfo()
        //                {
        //                    Property = property,
        //                    DefaultValue = value != null ? value.Value : null
        //                };
        //            }
        //        }
        //    }
        //    return null;
        //}
    }

    public class ExcelError
    {
        public int linha { get; set; }
        public int coluna { get; set; }
        public Exception excecao { get; set; }
    }
}
