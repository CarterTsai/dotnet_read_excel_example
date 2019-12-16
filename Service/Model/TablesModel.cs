namespace dotnet_read_excel_example.Service.Model
{
    public class TablesModel
    {
        /// <summary>
        /// 主鍵
        /// </summary>
        public string Key {get; set;}
        /// <summary>
        /// 欄位
        /// </summary>
        public string Column {get; set;}
        /// <summary>
        /// 資料型態 
        /// </summary>
        public string DataType {get; set;}
        /// <summary>
        /// Null 
        /// </summary>
        public string IsNull {get; set;}
        /// <summary>
        /// 欄位名稱 
        /// </summary>
        public string ColumnName {get; set;}
        /// <summary>
        /// 預設值 
        /// </summary>
        public string DefaultValue {get; set;}
        /// <summary>
        /// 備註 
        /// </summary>
        public string Comment {get; set;}

    }
}