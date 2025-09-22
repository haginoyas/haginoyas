/**
 * カスタム除外辞書を取得する
 * @returns {Object} カスタム除外辞書
 */
function getCustomExclusions() {
  try {
    const customExclusionsJson = PropertiesService.getScriptProperties().getProperty('CUSTOM_EXCLUSIONS');
    return customExclusionsJson ? JSON.parse(customExclusionsJson) : {};
  } catch (error) {
    console.log('カスタム除外辞書の読み込みエラー:', error);
    return {};
  }
}

/**
 * カスタム除外辞書を保存する
 * @param {Object} exclusions カスタム除外辞書
 */
function saveCustomExclusions(exclusions) {
  try {
    PropertiesService.getScriptProperties().setProperty('CUSTOM_EXCLUSIONS', JSON.stringify(exclusions));
  } catch (error) {
    console.log('カスタム除外辞書の保存エラー:', error);
  }
}

/**
 * 固定翻訳辞書を取得する
 * @returns {Object} 固定翻訳辞書
 */
function getFixedTranslations() {
  try {
    const fixedTranslationsJson = PropertiesService.getScriptProperties().getProperty('FIXED_TRANSLATIONS');
    return fixedTranslationsJson ? JSON.parse(fixedTranslationsJson) : {};
  } catch (error) {
    console.log('固定翻訳辞書の読み込みエラー:', error);
    return {};
  }
}

/**
 * 固定翻訳辞書を保存する
 * @param {Object} translations 固定翻訳辞書
 */
function saveFixedTranslations(translations) {
  try {
    PropertiesService.getScriptProperties().setProperty('FIXED_TRANSLATIONS', JSON.stringify(translations));
  } catch (error) {
    console.log('固定翻訳辞書の保存エラー:', error);
  }
}

/**
 * 翻訳除外辞書を管理する
 */
function manageTranslationExclusions() {
  const ui = SlidesApp.getUi();
  
  const response = ui.alert(
    '翻訳除外辞書管理',
    '翻訳から除外する単語を管理します。何を行いますか？',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    // 新しい単語を追加
    addCustomExclusion();
  } else if (response === ui.Button.NO) {
    // 現在の除外辞書を表示
    showCurrentExclusions();
  }
  // CANCELの場合は何もしない
}

/**
 * 固定翻訳辞書を管理する
 */
function manageFixedTranslations() {
  const ui = SlidesApp.getUi();
  
  const response = ui.alert(
    '固定翻訳辞書管理',
    '特定の単語を常に指定した翻訳に置き換える設定を管理します。何を行いますか？\n\nYes: 新しい固定翻訳を追加\nNo: 現在の固定翻訳一覧を表示\nCancel: 戻る',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    // 新しい固定翻訳を追加
    addFixedTranslation();
  } else if (response === ui.Button.NO) {
    // 現在の固定翻訳辞書を表示
    showCurrentFixedTranslations();
  }
  // CANCELの場合は何もしない
}

/**
 * カスタム除外単語を追加する
 */
function addCustomExclusion() {
  const ui = SlidesApp.getUi();
  
  const response = ui.prompt(
    '新しい除外単語を追加',
    '翻訳から除外したい単語を入力してください:\n例: MyService, CustomTool, SpecialAPI',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newTerm = response.getResponseText().trim();
    if (newTerm) {
      const customExclusions = getCustomExclusions();
      customExclusions[newTerm] = newTerm; // キーと値を同じにする
      saveCustomExclusions(customExclusions);
      
      ui.alert(`「${newTerm}」を翻訳除外辞書に追加しました！`);
    } else {
      ui.alert('単語が入力されていません。');
    }
  }
}

/**
 * 固定翻訳を追加する
 */
function addFixedTranslation() {
  const ui = SlidesApp.getUi();
  
  // 元の単語を入力
  const sourceResponse = ui.prompt(
    '固定翻訳追加 (1/2)',
    '元の単語（翻訳元）を入力してください:\n例: Hello, Good morning, Thank you',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sourceResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const sourceTerm = sourceResponse.getResponseText().trim();
  if (!sourceTerm) {
    ui.alert('元の単語が入力されていません。');
    return;
  }
  
  // 翻訳先の単語を入力
  const targetResponse = ui.prompt(
    '固定翻訳追加 (2/2)',
    `「${sourceTerm}」の翻訳先を入力してください:\n例: こんにちは, おはようございます, ありがとうございます`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (targetResponse.getSelectedButton() === ui.Button.OK) {
    const targetTerm = targetResponse.getResponseText().trim();
    if (targetTerm) {
      const fixedTranslations = getFixedTranslations();
      fixedTranslations[sourceTerm] = targetTerm;
      saveFixedTranslations(fixedTranslations);
      
      ui.alert(`固定翻訳を追加しました！\n「${sourceTerm}」→「${targetTerm}」`);
    } else {
      ui.alert('翻訳先が入力されていません。');
    }
  }
}

/**
 * 現在の除外辞書を表示する
 */
function showCurrentExclusions() {
  const ui = SlidesApp.getUi();
  const customExclusions = getCustomExclusions();
  const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };
  
  // デフォルト辞書とカスタム辞書を分けて表示
  const defaultTerms = Object.keys(TRANSLATION_EXCLUSIONS);
  const customTerms = Object.keys(customExclusions);
  
  let message = '【翻訳除外辞書】\n\n';
  
  message += `■ デフォルト除外用語 (${defaultTerms.length}個):\n`;
  message += defaultTerms.slice(0, 20).join(', '); // 最初の20個のみ表示
  if (defaultTerms.length > 20) {
    message += `\n...他 ${defaultTerms.length - 20}個`;
  }
  
  if (customTerms.length > 0) {
    message += `\n\n■ カスタム除外用語 (${customTerms.length}個):\n`;
    message += customTerms.join(', ');
    message += '\n\n※カスタム用語を削除する場合は、スクリプトエディタで手動削除してください。';
  } else {
    message += '\n\n■ カスタム除外用語: なし';
  }
  
  ui.alert('翻訳除外辞書', message, ui.ButtonSet.OK);
}

/**
 * 現在の固定翻訳辞書を表示する
 */
function showCurrentFixedTranslations() {
  const ui = SlidesApp.getUi();
  const fixedTranslations = getFixedTranslations();
  
  let message = '【固定翻訳辞書】\n\n';
  
  const entries = Object.entries(fixedTranslations);
  if (entries.length > 0) {
    message += `■ 固定翻訳 (${entries.length}個):\n`;
    entries.forEach(([source, target]) => {
      message += `「${source}」→「${target}」\n`;
    });
    message += '\n※固定翻訳を削除する場合は、「固定翻訳を削除」メニューを使用してください。';
  } else {
    message += '■ 固定翻訳: 未設定';
  }
  
  ui.alert('固定翻訳辞書', message, ui.ButtonSet.OK);
}

/**
 * 固定翻訳を削除する（管理機能）
 */
function removeFixedTranslation() {
  const ui = SlidesApp.getUi();
  const fixedTranslations = getFixedTranslations();
  const entries = Object.entries(fixedTranslations);
  
  if (entries.length === 0) {
    ui.alert('削除可能な固定翻訳がありません。');
    return;
  }
  
  // 削除する翻訳を選択（簡易版：単語入力）
  const response = ui.prompt(
    '固定翻訳削除',
    '削除したい元の単語を入力してください:\n\n現在の固定翻訳:\n' + 
    entries.map(([source, target]) => `「${source}」→「${target}」`).join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const termToRemove = response.getResponseText().trim();
    if (termToRemove && fixedTranslations[termToRemove]) {
      delete fixedTranslations[termToRemove];
      saveFixedTranslations(fixedTranslations);
      ui.alert(`「${termToRemove}」の固定翻訳を削除しました。`);
    } else {
      ui.alert('指定された単語が見つかりません。');
    }
  }
}

/**
 * 全てのスライドの日本語を校正する（基本）
 */
function proofreadAllSlidesJapanese() {
  proofreadAllSlides();
}

/**
 * 特定のスライドページの日本語を校正する（基本）
 */
function proofreadSpecificPageJapanese() {
  proofreadSpecificPage();
}

// Databricks API設定（スクリプトプロパティに設定してください）
const DATABRICKS_TOKEN = PropertiesService.getScriptProperties().getProperty('DATABRICKS_TOKEN');
const DATABRICKS_WORKSPACE = PropertiesService.getScriptProperties().getProperty('DATABRICKS_WORKSPACE');
const DATABRICKS_MODEL_ENDPOINT = 'databricks-meta-llama-3-1-70b-instruct'; // 無料で使用可能

/**
 * 翻訳除外辞書 - 翻訳しない単語・用語
 * 随時追加・編集可能
 */
const TRANSLATION_EXCLUSIONS = {
// 追加
'permission':'権限',
'Authentication':'Authentications',
'authorization':'認可',
'Tick Data':'Tick Data',
'Refined':'Refined',
'Raw': 'Raw',
'Aggregated': 'Aggregated',
'Query and Process': 'Query and Process',
'Transform': 'Transform',
'Ingest': 'Ingest',
'Delta Lake': 'Delta Lake',
'Serve': 'Serve',
'Platinum layer': 'Platinum layer',
'Analysis/Output': 'Analysis/Output',
// Databricks 公式
'4X-Large': '4X-Large',
'ACL': 'ACL',
'ACID': 'ACID',
'2X-Small': '2X-Small',
'2X-Large': '2X-Large',
'ADLS': 'ADLS',
'3X-Large': '3X-Large',
'ACR': 'ACR',
'Apache Kudu': 'Apache Kudu',
'Apache Spark™': 'Apache Spark™',
'Apache Software Foundation': 'Apache Software Foundation',
'Apache': 'Apache',
'ARNs': 'ARNs',
'APIs': 'APIs',
'ASF': 'ASF',
'Apache Hive': 'Apache Hive',
'Auto Loader': 'Auto Loader',
'AWS Marketplace': 'AWS Marketplace',
'Azure Data Lake Storage Gen2': 'Azure Data Lake Storage Gen2',
'Apache Kylin': 'Apache Kylin',
'Apache Spark': 'Apache Spark',
'Avro': 'Avro',
'AVRO': 'AVRO',
'Azure': 'Azure',
'AWS': 'AWS',
'Azure Synapse Analytics': 'Azure Synapse Analytics',
'Azure SQL Data Warehouse': 'Azure SQL Data Warehouse',
'Azure Databricks': 'Azure Databricks',
'Bayesian': 'Bayesian',
'BINARYFILE': 'BINARYFILE',
'Azure Key Vault': 'Azure Key Vault',
'Azure Blob': 'Azure Blob',
'Azure Event Hub': 'Azure Event Hub',
'BigQuery': 'BigQuery',
'Bokeh': 'Bokeh',
'Cassandra': 'Cassandra',
'CDC': 'CDC',
'Boxplot': 'Boxplot',
'Bitbucket': 'Bitbucket',
'CIDR': 'CIDR',
'Azure Cosmos DB': 'Azure Cosmos DB',
'ARN': 'ARN',
'Azure Data Lake Storage Gen1': 'Azure Data Lake Storage Gen1',
'Azure Active Directory': 'Azure Active Directory',
'API': 'API',
'Azure Synapse': 'Azure Synapse',
'Berkeley Packet Filter': 'Berkeley Packet Filter',
'Conda': 'Conda',
'BPF': 'BPF',
'CI/CD': 'CI/CD',
'CNN': 'CNN',
'Azure Data Lake Store': 'Azure Data Lake Store',
'CRAN': 'CRAN',
'Community Edition': 'Community Edition',
'Couchbase': 'Couchbase',
'Data Explorer': 'Data Explorer',
'data source': 'data source',
'CSV': 'CSV',
'Databricks Community Edition': 'Databricks Community Edition',
'Databricks Delta': 'Databricks Delta',
'Databricks Filesystem': 'Databricks Filesystem',
'Databricks DataInsight': 'Databricks DataInsight',
'Databricks': 'Databricks',
'Databricks Machine Learning': 'Databricks Machine Learning',
'Databricks Data Analytics Platform': 'Databricks Data Analytics Platform',
'DataFrames': 'DataFrames',
'DBU': 'DBU',
'Databricks Runtime': 'Databricks Runtime',
'Databricks Platform Free Trial': 'Databricks Platform Free Trial',
'Databricks Container Services': 'Databricks Container Services',
'DataFrame': 'DataFrame',
'data scientists': 'data scientists',
'Databricks SQL': 'Databricks SQL',
'Delta': 'Delta',
'Delta Lake': 'Delta Lake',
'Delta Sharing': 'Delta Sharing',
'Databricks Workflows': 'Databricks Workflows',
'DBFS': 'DBFS',
'DevOps': 'DevOps',
'DLT': 'DLT',
'Donut': 'Donut',
'Docker': 'Docker',
'Docker Container Services': 'Docker Container Services',
'eBay': 'eBay',
'Delta Engine': 'Delta Engine',
'EC2': 'EC2',
'Amazon ECR': 'Amazon ECR',
'Databricks Workspace': 'Databricks Workspace',
'EMR': 'EMR',
'Elasticsearch': 'Elasticsearch',
'ETL': 'ETL',
'GCS': 'GCS',
'Excel': 'Excel',
'GATK': 'GATK',
'Google Cloud Storage': 'Google Cloud Storage',
'GitHub': 'GitHub',
'Hadoop': 'Hadoop',
'HIPAA': 'HIPAA',
'Hive': 'Hive',
'eQTL': 'eQTL',
'GovCloud': 'GovCloud',
'Git': 'Git',
'Google Spreadsheets': 'Google Spreadsheets',
'Hive SerDes': 'Hive SerDes',
'HMS': 'HMS',
'Horovod': 'Horovod',
'Hugging Face': 'Hugging Face',
'IAM': 'IAM',
'HP': 'HP',
'IoT': 'IoT',
'Java': 'Java',
'JAR': 'JAR',
'iPass': 'iPass',
'JVM': 'JVM',
'JSON': 'JSON',
'Kafka': 'Kafka',
'GraphX': 'GraphX',
'HDFS': 'HDFS',
'Fivetran': 'Fivetran',
'GCP': 'GCP',
'Lambda': 'Lambda',
'Kinesis': 'Kinesis',
'LZO': 'LZO',
'MapReduce': 'MapReduce',
'HomeAway': 'HomeAway',
'Keras': 'Keras',
'Hive metastore': 'Hive metastore',
'IT': 'IT',
'Looker': 'Looker',
'MLlib': 'MLlib',
'Maven': 'Maven',
'Koalas': 'Koalas',
'KMS': 'KMS',
'MongoDB': 'MongoDB',
'MongoDB Connector for Spark': 'MongoDB Connector for Spark',
'Microsoft': 'Microsoft',
'MLOps': 'MLOps',
'Netflix': 'Netflix',
'NoSQL': 'NoSQL',
'Nephos': 'Nephos',
'NIST': 'NIST',
'MongoDB Atlas': 'MongoDB Atlas',
'NumPy': 'NumPy',
'MySQL': 'MySQL',
'ODBC': 'ODBC',
'Matplotlib': 'Matplotlib',
'MLflow': 'MLflow',
'ORC': 'ORC',
'Pandas': 'Pandas',
'Partner Connect': 'Partner Connect',
'Partner Hub': 'Partner Hub',
'Parquet': 'Parquet',
'PARQUET': 'PARQUET',
'Pandas DataFrame': 'Pandas DataFrame',
'Photon': 'Photon',
'PHI': 'PHI',
'PR': 'PR',
'PostgreSQL': 'PostgreSQL',
'PySpark': 'PySpark',
'PyPI': 'PyPI',
'PyCharm': 'PyCharm',
'Qlik Sense': 'Qlik Sense',
'Python': 'Python',
'R': 'R',
'PyTorch': 'PyTorch',
'Regeneron': 'Regeneron',
'Redshift': 'Redshift',
'ROC': 'ROC',
'RStudio': 'RStudio',
'repo': 'repo',
'SAML': 'SAML',
'SCD': 'SCD',
'Plotly': 'Plotly',
'RDBMS': 'RDBMS',
'RDD': 'RDD',
'SerDes': 'SerDes',
'Scala': 'Scala',
'S3': 'S3',
'Redash': 'Redash',
'Runtime': 'Runtime',
'REST': 'REST',
'SCIM': 'SCIM',
'Spark UI': 'Spark UI',
'Spark SQL': 'Spark SQL',
'SparkHub': 'SparkHub',
'Spark': 'Spark',
'Scikit-Learn': 'Scikit-Learn',
'SLA': 'SLA',
'SparkR': 'SparkR',
'SQL': 'SQL',
'Sparklyr': 'Sparklyr',
'SSD': 'SSD',
'Repos': 'Repos',
'SaaS': 'SaaS',
'Tableau': 'Tableau',
'r4.4large': 'r4.4large',
'Synapse': 'Synapse',
'Terraform': 'Terraform',
'Snowflake': 'Snowflake',
'Spark API': 'Spark API',
'TensorBoard': 'TensorBoard',
'TensorFlow': 'TensorFlow',
'TEXT': 'TEXT',
'Tungsten': 'Tungsten',
'unified analytics': 'unified analytics',
'Unity Catalog': 'Unity Catalog',
'Unified Data Analytics': 'Unified Data Analytics',
'X-Small': 'X-Small',
'XGBoost': 'XGBoost',
'VPC': 'VPC',
'Yahoo': 'Yahoo',
'X-Large': 'X-Large',
'Kudu': 'Kudu',
'Kylin': 'Kylin',
'Model Registry': 'Model Registry',
'azure': 'azure',
'Workspace Model Registry': 'Workspace Model Registry',
'MLflow Model Registry': 'MLflow Model Registry',
'Databricks Marketplace': 'Databricks Marketplace',
'SQL Server': 'SQL Server',
'CLI': 'CLI',
'Resilient Distributed Datasets': 'Resilient Distributed Datasets',
'write.parquet': 'write.parquet',
'write.orc': 'write.orc',
'write.text': 'write.text',
'read.json': 'read.json',
'read.parquet': 'read.parquet',
'read.orc': 'read.orc',
'read.text': 'read.text',
'read.jdbc': 'read.jdbc',
'change data feed': 'change data feed',
'external location': 'external location',
'Low-Med': 'Low-Med',
'Med-High': 'Med-High',
'CAN ATTACH TO': 'CAN ATTACH TO',
'CAN RESTART': 'CAN RESTART',
'production': 'production',
'warehouse': 'warehouse',
'Amazon': 'Amazon',
'Airflow': 'Airflow',
'Databricks Repo': 'Databricks Repo',
'SDK': 'SDK',
'SBT': 'SBT',
'DHCP': 'DHCP',
'FSCK': 'FSCK',
'PowerShell': 'PowerShell',
'zsh': 'zsh',
'IaC': 'IaC',
'Msck': 'Msck',
'TSV': 'TSV',
'SSL': 'SSL',
'UID': 'UID',
'PWD': 'PWD',
'UDF': 'UDF',
'rh': 'rh',
'ML': 'ML',
'LTS': 'LTS',
'write.json': 'write.json',
'COPY INTO': 'COPY INTO',
'Lakehouse Monitoring': 'Lakehouse Monitoring',
'Petastorm': 'Petastorm',
'Hyperopt': 'Hyperopt',
'Prophet': 'Prophet',
'LightGBM': 'LightGBM',
'ARIMA': 'ARIMA',
'Auto-ARIMA': 'Auto-ARIMA',
'Shapley': 'Shapley',
'sklearn': 'sklearn',
'scikit-learn': 'scikit-learn',
'PyTorch Lightning': 'PyTorch Lightning',
'HorovodRunner': 'HorovodRunner',
'TorchDistributor': 'TorchDistributor',
'John Snow Labs': 'John Snow Labs',
'BERT': 'BERT',
'LangChain': 'LangChain',
'ALTER CATALOG': 'ALTER CATALOG',
'ALTER CREDENTIAL': 'ALTER CREDENTIAL',
'ALTER DATABASE': 'ALTER DATABASE',
'ALTER LOCATION': 'ALTER LOCATION',
'ALTER PROVIDER': 'ALTER PROVIDER',
'ALTER RECIPIENT': 'ALTER RECIPIENT',
'ALTER TABLE': 'ALTER TABLE',
'lineage': 'lineage',
'customer-managed VPC': 'customer-managed VPC',
'billable usage log': 'billable usage log',
'billable usage': 'billable usage',
'Boto': 'Boto',
'Z-order': 'Z-order',
'Z-ordering': 'Z-ordering',
'source': 'source',
'ALTER SCHEMA': 'ALTER SCHEMA',
'ALTER SHARE': 'ALTER SHARE',
'ALTER VIEW': 'ALTER VIEW',
'Well-Architected': 'Well-Architected',
'Google Cloud Architecture Framework': 'Google Cloud Architecture Framework',
'COMMENT ON': 'COMMENT ON',
'COMMENT ON PROVIDER': 'COMMENT ON PROVIDER',
'COMMENT ON SHARE': 'COMMENT ON SHARE',
'COMMENT ON RECIPIENT': 'COMMENT ON RECIPIENT',
'COMMENT ON VOLUME': 'COMMENT ON VOLUME',
'COMMENT ON CONNECTION': 'COMMENT ON CONNECTION',
'CREATE BLOOMFILTER INDEX': 'CREATE BLOOMFILTER INDEX',
'CREATE CATALOG': 'CREATE CATALOG',
'CREATE FUNCTION': 'CREATE FUNCTION',
'CREATE LOCATION': 'CREATE LOCATION',
'CREATE RECIPIENT': 'CREATE RECIPIENT',
'CREATE SCHEMA': 'CREATE SCHEMA',
'CREATE TABLE CLONE': 'CREATE TABLE CLONE',
'liquid clustering': 'liquid clustering',
'CREATE SHARE': 'CREATE SHARE',
'Iceberg': 'Iceberg',
'disaster recovery': 'disaster recovery',
'CREATE TABLE': 'CREATE TABLE',
'CREATE TABLE LIKE': 'CREATE TABLE LIKE',
'SHOW CREATE TABLE': 'SHOW CREATE TABLE',
'REPLACE TABLE': 'REPLACE TABLE',
'DESC': 'DESC',
'Lakeview': 'Lakeview',
'Network Connectivity Configs': 'Network Connectivity Configs',
'SparkSession': 'SparkSession',
'GitHub Actions': 'GitHub Actions',
'Snowplow': 'Snowplow',
'Arcion': 'Arcion',
'Hevo': 'Hevo',
'Hevo Data': 'Hevo Data',
'Infoworks': 'Infoworks',
'DataFoundry': 'DataFoundry',
'Qlik': 'Qlik',
'Qlik Replicate': 'Qlik Replicate',
'Qlik Compose': 'Qlik Compose',
'Rivery': 'Rivery',
'Stitch': 'Stitch',
'StreamSets': 'StreamSets',
'Syncsort': 'Syncsort',
'dbt': 'dbt',
'dbt Cloud': 'dbt Cloud',
'dbt Core': 'dbt Core',
'Matillion': 'Matillion',
'Prophecy': 'Prophecy',
'Labelbox': 'Labelbox',
'Hex': 'Hex',
'MicroStrategy': 'MicroStrategy',
'Mode': 'Mode',
'Sigma': 'Sigma',
'SQL Workbench/J': 'SQL Workbench/J',
'ThoughtSpot': 'ThoughtSpot',
'TIBCO Spotfire': 'TIBCO Spotfire',
'TIBCO Spotfire Analyst': 'TIBCO Spotfire Analyst',
'Census': 'Census',
'Hightouch': 'Hightouch',
'AtScale': 'AtScale',
'Stardog': 'Stardog',
'Alation': 'Alation',
'Anomalo': 'Anomalo',
'erwin Data Modeler': 'erwin Data Modeler',
'erwin Data Modeler by Quest': 'erwin Data Modeler by Quest',
'Lightup': 'Lightup',
'Hunters': 'Hunters',
'Privacera': 'Privacera',
'CREATE VIEW': 'CREATE VIEW',
'Microsoft Entra ID': 'Microsoft Entra ID',
'Databricks Connect': 'Databricks Connect',
'Windows': 'Windows',
'RudderStack': 'RudderStack',
'WindowSpecDefinition': 'WindowSpecDefinition',
'2.0.61.Final-windows-x86_64': '2.0.61.Final-windows-x86_64',
'feature table': 'feature table',
'deletion vector': 'deletion vector',
'Vector Search': 'Vector Search',
'Z-ordered': 'Z-ordered',
'Poetry': 'Poetry',
'Generative AI': 'Generative AI',
'large language model': 'large language model',
'LLM': 'LLM',
'foundation model': 'foundation model',
'fine tuning': 'fine tuning',
'AI Functions': 'AI Functions',
'Transformer': 'Transformer',
'DetrForSegmentation': 'DetrForSegmentation',
'pyodbc': 'pyodbc',
'TABLESAMPLE': 'TABLESAMPLE',
'Node.js': 'Node.js',
'Node Version Manager': 'Node Version Manager',
'CDKTF CLI': 'CDKTF CLI',
'Graviton': 'Graviton',
'UDAFs': 'UDAFs',
'globber': 'globber',
'AI Playground': 'AI Playground',
'Genie': 'Genie',
'AI Functions on Databricks': 'AI Functions on Databricks',
'Eclipse': 'Eclipse',
'IntelliJ IDEA': 'IntelliJ
