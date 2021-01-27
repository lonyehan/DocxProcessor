# DocxProcessor
## 功能介紹
  ### 取代功能
  1. 用字串取代字串
  2. 用圖片取代字串
  3. 用ModelList取代字串(插入多列)
  ### 合併功能
  1. 將多個Docx合併
  
## 使用說明
### 取代功能
  1. 用字串字典樹取代字串的部分
  ```C#
    using DocxProcessor
    
    // Filepath as input, byte array as output    
    
    // Template File Path 
    string TemplatePath = "<TargetFile>.docx";  
            
    // prepare replace items 
    Dictiondary<string, string> items = new Dictiondary<string, string>();    
    
    items.Add("#SearchingString#", "ReplaceString")    
    ...
    
    // new Replacer
    var Replacer = new ReplaceWordTemplate();
    
    // Replace 
    Replacer.Replace(TemplatePath, items);         
 ```
 
 2. 用Model來取代字串的部分
  在這邊會將Model的欄位名稱前後加上##方便替代<br/>
  例如以下提及的No，實質上替代的會是檔案中的#No#
 ```C#
    using DocxProcessor
    
    // Filepath as input, byte array as output    
    
    // Template File Path 
    string TemplatePath = "<TargetFile>.docx";
            
    // prepare replace Model
    CustomModel Model = new CustomModel();
    Model.No = 200;    
    ...
    
    // new Replacer
    var Replacer = new ReplaceWordTemplate();
    
    // Replace 
    Replacer.Replace(TemplatePath, Model);         
 ```
 
 3. 用圖片代替字串(目前在欄位內測試沒有問題)
 ```C#
    using DocxProcessor
    
    // Filepath as input, byte array as output    
    
    // Template File Path 
    string TemplatePath = "<TargetFile>.docx";
                
    // prepare replace Image
    // ImagePath
    string imagePath = "<TargetImage>.png" // png or jpg or gif
    
    Dictionary<string, ImageData> items = new Dictionary<string, ImageData>();
    
    ImageData image = new ImageData(imagePath, Width: <customWidth>, Height: <customHeight>); // Unit is centimeter(CM)
    
    items.Add("SearchingString", image);        
    ...
    
    // new Replacer
    var Replacer = new ReplaceWordTemplate();
    
    // Replace 
    Replacer.ReplaceTableCellByImage(TemplatePath, items);         
 ```
 
 4. 用List<Model>的方式取代字串
 這邊的使用會是帶入資料列去新增行列
  ```C#
    using DocxProcessor
    
    // Filepath as input, byte array as output    
    
    // Template File Path 
    string TemplatePath = "<TargetFile>.docx";
            
    // prepare replace ModelList
    List<CustomModel> ModelList = new List<CustomModel>();
    
    CustomModel customModel1 = new CustomModel();
    CustomModel customModel2 = new CustomModel();
    
    ModelList.Add(customModel1);
    ModelList.Add(customModel2);    
    ...
    
    // new Replacer
    var Replacer = new ReplaceWordTemplate();
    
    // Replace 
    Replacer.Replace(TemplatePath, ModelList);         
 ```
 
 ## 合併功能
  1. 將多筆Docx合併到同一份資料
  ```C#
    // Filepath as input, byte array as output    
    
    // Docxs Filepath 
    string Filepath1 = "<TargetFile1>.docx";
    string Filepath2 = "<TargetFile2>.docx";
    
    // prepare docxs to merge
    // merge these docxs
    List<Stream> docxs = new List<Stream>();
    
    // merge
    FileStream docx1 = new FileStream(Filepath1, FileMode.Open);
    FileStream docx2 = new FileStream(Filepath2, FileMode.Open);
    
    docxs.Add(docx1);
    docxs.Add(docx2);       
    ...
    
    // new Merger
    var Merger = new MergeWordTemplate();
    
    // Replace 
    Merger.MergeDocxsInotOne(docxs);
    
  ```
