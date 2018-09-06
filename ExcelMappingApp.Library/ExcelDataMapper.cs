namespace ExcelMappingApp.Library
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq.Expressions;

    public class ExcelDataMapper<TModel> where TModel : IExcelMappingModel
    {
        private Stream _stream;

        private List<ColumnMapConfiguration> ColumnMapConfigurations { get; set; } = new List<ColumnMapConfiguration>();
        private List<string> Errors { get; set; } = new List<string>();

        public ExcelDataMapper<TModel> Bind(Expression<Func<TModel, object>> property, string headerColumnAddress, string headerColumnTitle, Func<object, object> converter = null)
        {
            ExcelUtilities.EnsureColumnValid(headerColumnAddress);

            var propertyName = string.Empty;

            try
            {
                propertyName = ((MemberExpression)property.Body).Member.Name;
            }
            catch
            {
                propertyName = ((MemberExpression)((UnaryExpression)property.Body).Operand).Member.Name;
            }

            ColumnMapConfigurations.Add(new ColumnMapConfiguration
            {
                PropertyName = propertyName,
                HeaderAddress = headerColumnAddress,
                HeaderTitle = headerColumnTitle,
                Converter = converter
            });

            return this;
        }

        public ExcelDataMapper<TModel> FromStream(Stream stream)
        {
            EnsureValidStream(stream);

            _stream = stream;

            return this;
        }

        public List<TModel> Map()
        {
            var result = new List<TModel>();

            EnsureValidStream(_stream);

            using (var excelPackage = new ExcelPackage(_stream))
            {
                var workSheet = excelPackage.Workbook.Worksheets[0];

                var typeModel = typeof(TModel);

                foreach (var configuration in ColumnMapConfigurations)
                {
                    var propertyInfo = typeModel.GetProperty(configuration.PropertyName);
                    var excelValue = workSheet.Cells[configuration.HeaderAddress].Value;

                    EnsureValidHeader(configuration.HeaderTitle, excelValue?.ToString());
                }

                if (Errors.Count > 0)
                    throw new NotMatchHeaderException(string.Join("\n", Errors));
                
                for (int i = 2; i <= workSheet.Dimension.End.Row; i++)
                {
                    var model = (IExcelMappingModel)Activator.CreateInstance(typeModel);
                    model.Row = i.ToString();

                    foreach (var configuration in ColumnMapConfigurations)
                    {
                        var propertyInfo = typeModel.GetProperty(configuration.PropertyName);
                        var columnParts = ExcelUtilities.GetPartsOfColumn(configuration.HeaderAddress);

                        var excelValue = GetConvertedValue(workSheet.Cells, columnParts.Item1 + i, propertyInfo.PropertyType, configuration.Converter);

                        propertyInfo.SetValue(model, excelValue, null);
                    }

                    result.Add((TModel)model);
                }

                if (Errors.Count > 0)
                    throw new UnexpectedValueTypeException(string.Join("\n", Errors));
            }

            return result;
        }

        #region Helpers

        private void EnsureValidStream(Stream stream)
        {
            if (stream == null || stream.Length == 0)
                throw new ArgumentException("stream");
        }

        private void EnsureValidHeader(string expect, string actual)
        {
            if (expect != actual)
                Errors.Add(string.Format("Beklenen Başlık Değeri : {0} - Gelen Başlık Değeri : {1}", expect, actual));
        }

        private object GetConvertedValue(ExcelRange excelRange, string address, Type converterType, Func<object, object> customConverter = null)
        {
            var excelValue = excelRange[address].Value;

            try
            {
                return customConverter == null ? Convert.ChangeType(excelValue, converterType) : customConverter(excelValue);
            }
            catch
            {
                Errors.Add(string.Format("{0} kolonunda uyumsuz değer var. Gelen Değer: {1}", address, excelValue));
            }

            return null;
        }

        #endregion
    }
}
