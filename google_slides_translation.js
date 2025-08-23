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
  // Databricks プラットフォーム・製品
  'Databricks': 'Databricks',
  'BI': 'BI',
  'Databricks Platform': 'Databricks Platform',
  'Databricks Lakehouse Platform': 'Databricks Lakehouse Platform',
  'Lakehouse': 'Lakehouse',
  'Data Lakehouse': 'Data Lakehouse',
  'Lakehouse Architecture': 'Lakehouse Architecture',
  
  // Databricks Unity Catalog関連
  'Unity Catalog': 'Unity Catalog',
  'UC': 'UC',
  'Unity Catalog Metastore': 'Unity Catalog Metastore',
  'Catalog': 'Catalog',
  'Schema': 'Schema',
  'External Location': 'External Location',
  'Storage Credential': 'Storage Credential',
  'Data Lineage': 'Data Lineage',
  'Data Governance': 'Data Governance',
  
  // Databricks Delta関連
  'Delta Lake': 'Delta Lake',
  'Delta': 'Delta',
  'Delta Table': 'Delta Table',
  'Delta Tables': 'Delta Tables',
  'Delta Sharing': 'Delta Sharing',
  'Delta Live Tables': 'Delta Live Tables',
  'DLT': 'DLT',
  'Delta Engine': 'Delta Engine',
  'Delta Cache': 'Delta Cache',
  'Delta Log': 'Delta Log',
  'Delta Merge': 'Delta Merge',
  'Delta Clone': 'Delta Clone',
  'Delta Time Travel': 'Delta Time Travel',
  'OPTIMIZE': 'OPTIMIZE',
  'VACUUM': 'VACUUM',
  'Z-ORDER': 'Z-ORDER',
  
  // Databricks MLflow関連
  'MLflow': 'MLflow',
  'MLflow Tracking': 'MLflow Tracking',
  'MLflow Projects': 'MLflow Projects',
  'MLflow Models': 'MLflow Models',
  'MLflow Registry': 'MLflow Registry',
  'Model Registry': 'Model Registry',
  'MLflow Serving': 'MLflow Serving',
  'MLflow Experiments': 'MLflow Experiments',
  'MLflow Runs': 'MLflow Runs',
  'MLflow Artifacts': 'MLflow Artifacts',
  
  // Databricks Compute関連
  'Compute': 'Compute',
  'Cluster': 'Cluster',
  'Clusters': 'Clusters',
  'All-Purpose Cluster': 'All-Purpose Cluster',
  'Job Cluster': 'Job Cluster',
  'SQL Warehouse': 'SQL Warehouse',
  'SQL Warehouses': 'SQL Warehouses',
  'Serverless': 'Serverless',
  'Serverless SQL Warehouse': 'Serverless SQL Warehouse',
  'Serverless Compute': 'Serverless Compute',
  'Driver Node': 'Driver Node',
  'Worker Node': 'Worker Node',
  'Autoscaling': 'Autoscaling',
  'Auto Termination': 'Auto Termination',
  'Spot Instance': 'Spot Instance',
  'Preemptible Instance': 'Preemptible Instance',
  'Pool': 'Pool',
  'Instance Pool': 'Instance Pool',
  
  // Databricks Photon関連
  'Photon': 'Photon',
  'Photon Engine': 'Photon Engine',
  'Photon-enabled': 'Photon-enabled',
  
  // Databricks Workflows関連
  'Workflows': 'Workflows',
  'Jobs': 'Jobs',
  'Job': 'Job',
  'Task': 'Task',
  'Multi-task Job': 'Multi-task Job',
  'Workflow Job': 'Workflow Job',
  'Scheduled Job': 'Scheduled Job',
  'Pipeline': 'Pipeline',
  'Data Pipeline': 'Data Pipeline',
  'ETL Pipeline': 'ETL Pipeline',
  
  // Databricks Notebook・開発環境関連
  'Notebook': 'Notebook',
  'Notebooks': 'Notebooks',
  'Databricks Notebook': 'Databricks Notebook',
  'Collaborative Notebook': 'Collaborative Notebook',
  'Repo': 'Repo',
  'Repos': 'Repos',
  'Git Integration': 'Git Integration',
  'Workspace': 'Workspace',
  'Workspace Files': 'Workspace Files',
  'DBFS': 'DBFS',
  'Databricks File System': 'Databricks File System',
  'Magic Command': 'Magic Command',
  'Cell': 'Cell',
  'Command': 'Command',
  
  // Databricks Auto Loader関連
  'Auto Loader': 'Auto Loader',
  'Autoloader': 'Autoloader',
  'Structured Streaming': 'Structured Streaming',
  'Streaming': 'Streaming',
  'Stream Processing': 'Stream Processing',
  'Incremental Processing': 'Incremental Processing',
  'Cloud Files': 'Cloud Files',
  'Trigger': 'Trigger',
  
  // Databricks Feature Store関連
  'Feature Store': 'Feature Store',
  'Feature Engineering': 'Feature Engineering',
  'Feature Table': 'Feature Table',
  'Feature Lookup': 'Feature Lookup',
  'Feature Serving': 'Feature Serving',
  'Feature Pipeline': 'Feature Pipeline',
  'Online Store': 'Online Store',
  'Offline Store': 'Offline Store',
  
  // Databricks SQL関連
  'Databricks SQL': 'Databricks SQL',
  'SQL Analytics': 'SQL Analytics',
  'SQL Editor': 'SQL Editor',
  'Query': 'Query',
  'Dashboard': 'Dashboard',
  'Dashboards': 'Dashboards',
  'Alert': 'Alert',
  'Alerts': 'Alerts',
  'Query Profile': 'Query Profile',
  'Query History': 'Query History',
  'SQL Endpoint': 'SQL Endpoint',
  'SQL Persona': 'SQL Persona',
  
  // Databricks Machine Learning関連
  'Databricks Machine Learning': 'Databricks Machine Learning',
  'Databricks ML': 'Databricks ML',
  'ML Runtime': 'ML Runtime',
  'MLR': 'MLR',
  'AutoML': 'AutoML',
  'Automated ML': 'Automated ML',
  'Hyperopt': 'Hyperopt',
  'Hyperparameter Tuning': 'Hyperparameter Tuning',
  'Distributed Training': 'Distributed Training',
  'Model Serving': 'Model Serving',
  'Model Endpoint': 'Model Endpoint',
  'Inference': 'Inference',
  'Real-time Inference': 'Real-time Inference',
  'Batch Inference': 'Batch Inference',
  
  // Databricks Partner Connect関連
  'Partner Connect': 'Partner Connect',
  'Partner Solutions': 'Partner Solutions',
  'Marketplace': 'Marketplace',
  'Solution Accelerator': 'Solution Accelerator',
  'Solution Accelerators': 'Solution Accelerators',
  
  // Databricks セキュリティ関連
  'Access Control': 'Access Control',
  'RBAC': 'RBAC',
  'Table Access Control': 'Table Access Control',
  'Column-level Security': 'Column-level Security',
  'Row-level Security': 'Row-level Security',
  'Audit Log': 'Audit Log',
  'Audit Logs': 'Audit Logs',
  'Secret Scope': 'Secret Scope',
  'Secrets': 'Secrets',
  'Personal Access Token': 'Personal Access Token',
  'PAT': 'PAT',
  'Service Principal': 'Service Principal',
  'SCIM': 'SCIM',
  'SSO': 'SSO',
  'SAML': 'SAML',
  'Identity Provider': 'Identity Provider',
  'VPC': 'VPC',
  'Private Link': 'Private Link',
  'Customer-managed Keys': 'Customer-managed Keys',
  'CMK': 'CMK',
  
  // Databricks API・SDK関連
  'Databricks API': 'Databricks API',
  'REST API': 'REST API',
  'Databricks SDK': 'Databricks SDK',
  'Databricks CLI': 'Databricks CLI',
  'Terraform Provider': 'Terraform Provider',
  'Databricks Connect': 'Databricks Connect',
  'ODBC': 'ODBC',
  'JDBC': 'JDBC',
  'Connector': 'Connector',
  'Driver': 'Driver',
  
  // Databricks Runtime関連
  'Databricks Runtime': 'Databricks Runtime',
  'DBR': 'DBR',
  'Runtime Version': 'Runtime Version',
  'LTS Runtime': 'LTS Runtime',
  'Databricks Runtime for ML': 'Databricks Runtime for ML',
  'Photon Runtime': 'Photon Runtime',
  'GPU Runtime': 'GPU Runtime',
  'Genomics Runtime': 'Genomics Runtime',
  
  // Databricks クラウド・リージョン関連
  'Databricks on AWS': 'Databricks on AWS',
  'Databricks on Azure': 'Databricks on Azure',
  'Databricks on GCP': 'Databricks on GCP',
  'Multi-cloud': 'Multi-cloud',
  'Cross-cloud': 'Cross-cloud',
  'Region': 'Region',
  'Availability Zone': 'Availability Zone',
  
  // Databricks 料金・プラン関連
  'Standard Tier': 'Standard Tier',
  'Premium Tier': 'Premium Tier',
  'Enterprise Tier': 'Enterprise Tier',
  'DBU': 'DBU',
  'Databricks Unit': 'Databricks Unit',
  'Consumption-based': 'Consumption-based',
  
  // Apache Spark関連（Databricksで重要）
  'Apache Spark': 'Apache Spark',
  'Spark': 'Spark',
  'Spark SQL': 'Spark SQL',
  'Spark Streaming': 'Spark Streaming',
  'Structured Streaming': 'Structured Streaming',
  'MLlib': 'MLlib',
  'GraphX': 'GraphX',
  'PySpark': 'PySpark',
  'Spark DataFrame': 'Spark DataFrame',
  'RDD': 'RDD',
  'Catalyst': 'Catalyst',
  'Tungsten': 'Tungsten',
  'Adaptive Query Execution': 'Adaptive Query Execution',
  'AQE': 'AQE',
  'Dynamic Partition Pruning': 'Dynamic Partition Pruning',
  'Broadcast Join': 'Broadcast Join',
  'Shuffle': 'Shuffle',
  'Partitioning': 'Partitioning',
  'Bucketing': 'Bucketing',
  
  // 一般IT用語
  'API': 'API',
  'SDK': 'SDK',
  'REST API': 'REST API',
  'GraphQL': 'GraphQL',
  'JSON': 'JSON',
  'XML': 'XML',
  'YAML': 'YAML',
  'SQL': 'SQL',
  'NoSQL': 'NoSQL',
  'ETL': 'ETL',
  'ELT': 'ELT',
  'CI/CD': 'CI/CD',
  'DevOps': 'DevOps',
  'MLOps': 'MLOps',
  'DataOps': 'DataOps',
  'Docker': 'Docker',
  'Kubernetes': 'Kubernetes',
  'Terraform': 'Terraform',
  'GitHub': 'GitHub',
  'GitLab': 'GitLab',
  'Azure': 'Azure',
  'AWS': 'AWS',
  'GCP': 'GCP',
  'Google Cloud': 'Google Cloud',
  
  // データ関連
  'Big Data': 'Big Data',
  'Data Lake': 'Data Lake',
  'Data Warehouse': 'Data Warehouse',
  'Data Pipeline': 'Data Pipeline',
  'ETL Pipeline': 'ETL Pipeline',
  'Real-time': 'Real-time',
  'Streaming': 'Streaming',
  'Batch Processing': 'Batch Processing',
  'Stream Processing': 'Stream Processing',
  'OLTP': 'OLTP',
  'OLAP': 'OLAP',
  
  // AI/ML用語
  'Machine Learning': 'Machine Learning',
  'Deep Learning': 'Deep Learning',
  'Neural Network': 'Neural Network',
  'Transformer': 'Transformer',
  'LLM': 'LLM',
  'Large Language Model': 'Large Language Model',
  'NLP': 'NLP',
  'Computer Vision': 'Computer Vision',
  'Feature Engineering': 'Feature Engineering',
  'Model Training': 'Model Training',
  'Inference': 'Inference',
  'AutoML': 'AutoML',
  'Hyperparameter': 'Hyperparameter',
  'Cross-validation': 'Cross-validation',
  
  // プログラミング言語・フレームワーク
  'Python': 'Python',
  'Scala': 'Scala',
  'Java': 'Java',
  'R': 'R',
  'JavaScript': 'JavaScript',
  'TypeScript': 'TypeScript',
  'React': 'React',
  'Vue.js': 'Vue.js',
  'Angular': 'Angular',
  'Node.js': 'Node.js',
  'Express': 'Express',
  'Flask': 'Flask',
  'Django': 'Django',
  'Spring Boot': 'Spring Boot',
  'TensorFlow': 'TensorFlow',
  'PyTorch': 'PyTorch',
  'Scikit-learn': 'Scikit-learn',
  'Pandas': 'Pandas',
  'NumPy': 'NumPy',
  
  // その他のサービス・ツール
  'Slack': 'Slack',
  'Microsoft Teams': 'Microsoft Teams',
  'Zoom': 'Zoom',
  'Jira': 'Jira',
  'Confluence': 'Confluence',
  'Tableau': 'Tableau',
  'Power BI': 'Power BI',
  'Power BI + Copilot': 'Power BI + Copilot',
  'Looker': 'Looker',
  'Snowflake': 'Snowflake',
  'Redshift': 'Redshift',
  'BigQuery': 'BigQuery',
  'MongoDB': 'MongoDB',
  'PostgreSQL': 'PostgreSQL',
  'MySQL': 'MySQL',
  'Redis': 'Redis',
  'Elasticsearch': 'Elasticsearch',
  
  // 追加のDatabricksプラットフォーム関連用語
  'Databricks Data Intelligence Platform': 'Databricks Data Intelligence Platform',
  'Data Intelligence Platform': 'Data Intelligence Platform',
  'Mosaic AI': 'Mosaic AI',
  'Databricks IQ': 'Databricks IQ',
  'Predictive IQ': 'Predictive IQ',
  'Predictive optimization': 'Predictive optimization',
  'Predictive Optimization': 'Predictive Optimization',
  'Assistant': 'Assistant',
  'AI Assistant': 'AI Assistant',
  
  // データレイヤー・階層
  'bronze': 'bronze',
  'silver': 'silver',
  'gold': 'gold',
  'Bronze': 'Bronze',
  'Silver': 'Silver',
  'Gold': 'Gold',
  'bronze layer': 'bronze layer',
  'silver layer': 'silver layer',
  'gold layer': 'gold layer',
  'Bronze Layer': 'Bronze Layer',
  'Silver Layer': 'Silver Layer',
  'Gold Layer': 'Gold Layer',
  
  // AI/BI関連
  'AI/BI': 'AI/BI',
  'AI/BI Dashboards': 'AI/BI Dashboards',
  'AI/BI Genie': 'AI/BI Genie',
  'Gen AI': 'Gen AI',
  'Generative AI': 'Generative AI',
  'Classic ML': 'Classic ML',
  'Vector Search': 'Vector Search',
  'AI Gateway': 'AI Gateway',
  'AI Functions': 'AI Functions',
  'AI Services': 'AI Services',
  
  // データ処理・エンジニアリング
  'Batch & Streaming': 'Batch & Streaming',
  'Data Engineering & Processing': 'Data Engineering & Processing',
  'Automation & Orchestration': 'Automation & Orchestration',
  'Workflows': 'Workflows',
  'CI/CD tools': 'CI/CD tools',
  'DataOps': 'DataOps',
  'MLOps': 'MLOps',
  
  // データ管理・ガバナンス
  'Data and AI Governance': 'Data and AI Governance',
  'Federation': 'Federation',
  'Catalog & Lineage': 'Catalog & Lineage',
  'Access Control': 'Access Control',
  'Data & AI Assets': 'Data & AI Assets',
  'Lakehouse Monitoring': 'Lakehouse Monitoring',
  'Data Management': 'Data Management',
  'Collaboration': 'Collaboration',
  'Clean Rooms': 'Clean Rooms',
  
  // ストレージ・データソース
  'ADLS Gen 2': 'ADLS Gen 2',
  'Azure Data Lake Storage Gen 2': 'Azure Data Lake Storage Gen 2',
  'Iceberg': 'Iceberg',
  'Apache Iceberg': 'Apache Iceberg',
  'HMS': 'HMS',
  'Hive Metastore': 'Hive Metastore',
  'IoT Hub': 'IoT Hub',
  'Event Hub': 'Event Hub',
  'Azure Data Factory': 'Azure Data Factory',
  'Fabric Data Factory': 'Fabric Data Factory',
  'Synapse SQL Server': 'Synapse SQL Server',
  'SQL DB': 'SQL DB',
  'Cosmos DB': 'Cosmos DB',
  'Operational DBs': 'Operational DBs',
  
  // 分析・可視化・アプリケーション
  'Query / Process': 'Query / Process',
  'Query/Process': 'Query/Process',
  'Data Analysis': 'Data Analysis',
  'Data Consumer': 'Data Consumer',
  'Business/AI App': 'Business/AI App',
  'Business AI App': 'Business AI App',
  'Sharing Partner': 'Sharing Partner',
  'Applications': 'Applications',
  'Databricks Apps': 'Databricks Apps',
  
  // Azure・Microsoft関連
  'Entra ID': 'Entra ID',
  'Azure Entra ID': 'Azure Entra ID',
  'ID Provider': 'ID Provider',
  'Purview': 'Purview',
  'Azure Purview': 'Azure Purview',
  'Copilot': 'Copilot',
  'Microsoft Copilot': 'Microsoft Copilot',
  
  // 外部サービス・プラットフォーム
  'Hugging Face': 'Hugging Face',
  'OpenAI': 'OpenAI',
  'External Orchestrator': 'External Orchestrator',
  'Connectors & APIs': 'Connectors & APIs',
  '3rd party': '3rd party',
  'Third party': 'Third party',
  
  // データ形式・構造
  'semi-structured': 'semi-structured',
  'unstructured': 'unstructured',
  'structured': 'structured',
  'Semi-structured': 'Semi-structured',
  'Unstructured': 'Unstructured',
  'Structured': 'Structured',
  'Files/Logs': 'Files/Logs',
  'Sensors and IoT': 'Sensors and IoT',
  'Business Apps': 'Business Apps',
  'Market places': 'Market places',
  'Data Shares': 'Data Shares'
}

/**
 * グーグルスライドの翻訳機能を追加するメニューを作成する
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('翻訳機能')
    .addItem('[全スライド] Google 翻訳（英語→日本語）', 'translateAllSlidesEnToJa')
    .addItem('[特定ページ] Google 翻訳（英語→日本語）', 'translateSpecificPageEnToJa')
    .addItem('[全スライド] Google 翻訳（日本語→英語）', 'translateAllSlidesJaToEn')
    .addItem('[特定ページ] Google 翻訳（日本語→英語）', 'translateSpecificPageJaToEn')
    .addSeparator()
    .addItem('[全スライド] 日本語を校正（基本）', 'proofreadAllSlidesJapanese')
    .addItem('[特定ページ] 日本語を校正（基本）', 'proofreadSpecificPageJapanese')
    .addSeparator()
    .addItem('[全スライド] Databricks AI校正', 'databricksProofreadAllSlides')
    .addItem('[特定ページ] Databricks AI校正', 'databricksProofreadSpecificPage')
    .addSeparator()
    .addItem('翻訳除外辞書を管理', 'manageTranslationExclusions')
    .addItem('Databricks API設定', 'setupDatabricksAPI')
    .addToUi();
}

/**
 * 全てのスライドをGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateAllSlidesEnToJa() {
  translateAllSlides('en', 'ja');
}

/**
 * 特定のスライドページをGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateSpecificPageEnToJa() {
  translateSpecificPage('en', 'ja');
}

/**
 * 全てのスライドをGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateAllSlidesJaToEn() {
  translateAllSlides('ja', 'en');
}

/**
 * 特定のスライドページをGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateSpecificPageJaToEn() {
  translateSpecificPage('ja', 'en');
}

/**
 * Databricks APIを使用してシェイプのテキストを校正する
 * @param {Shape} shape 校正対象のシェイプ
 */
function databricksProofreadShape(shape) {
  for (let paragraphText of shape.getText().getParagraphs()) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const proofreadText = proofreadWithDatabricks(originalText);
    paragraphText.setText(proofreadText);
  }
}

/**
 * Databricks APIを使用してグループ内のシェイプのテキストを校正する
 * @param {Group} group 校正対象のグループ
 */
function databricksProofreadGroup(group) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      databricksProofreadShape(childElement.asShape());
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      databricksProofreadGroup(childElement.asGroup());
    }
  }
}

/**
 * Databricks APIを使用してスライド内のテーブルを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function databricksProofreadTableInSlide(slide) {
  const tables = slide.getTables();

  for (let table of tables) {
    const numRows = table.getNumRows();
    const numColumns = table.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numColumns; col++) {
        try {
          const cell = table.getCell(row, col);
          const originalText = cell.getText().asString();
          if (originalText === '' || originalText.trim() === '') {
            continue;
          }

          const proofreadText = proofreadWithDatabricks(originalText);
          cell.getText().setText(proofreadText);
        } catch (error) {
          console.log(`セル (${row}, ${col}) をスキップしました: ${error.message}`);
          continue;
        }
      }
    }
  }
}

/**
 * Databricks APIを使用してスピーカーノートを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function databricksProofreadSpeakerNotes(slide) {
  const notesPageTexts = slide.getNotesPage().getSpeakerNotesShape().getText().getParagraphs();
  for(let paragraphText of notesPageTexts) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const proofreadText = proofreadWithDatabricks(originalText);
    paragraphText.setText(proofreadText);
  }
}

/**
 * Databricks API設定を行う
 */
function setupDatabricksAPI() {
  const ui = SlidesApp.getUi();
  
  // ワークスペースURLの設定
  const workspaceResponse = ui.prompt(
    'Databricks API設定 (1/2)',
    'Databricksワークスペースの名前を入力してください:\n例: your-workspace\n※ https://your-workspace.cloud.databricks.com のyour-workspace部分',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (workspaceResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const workspace = workspaceResponse.getResponseText().trim();
  if (!workspace) {
    ui.alert('ワークスペース名が入力されていません。');
    return;
  }
  
  // アクセストークンの設定
  const tokenResponse = ui.prompt(
    'Databricks API設定 (2/2)',
    'Databricksアクセストークンを入力してください:\n\n※トークンは安全に暗号化されて保存されます',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (tokenResponse.getSelectedButton() === ui.Button.OK) {
    const token = tokenResponse.getResponseText().trim();
    if (token) {
      PropertiesService.getScriptProperties().setProperties({
        'DATABRICKS_WORKSPACE': workspace,
        'DATABRICKS_TOKEN': token
      });
      ui.alert('Databricks API設定が完了しました！');
    } else {
      ui.alert('アクセストークンが入力されていません。');
    }
  }
}

/**
 * テキスト内の翻訳除外用語を保護し、翻訳後に復元する
 * @param {string} text 処理するテキスト
 * @param {function} translationFunction 翻訳処理を行う関数
 * @returns {string} 除外用語を保護した翻訳結果
 */
function protectAndTranslate(text, translationFunction) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return text;
  }

  let protectedText = text;
  const placeholders = {};
  let placeholderIndex = 0;

  // カスタム除外辞書を取得
  const customExclusions = getCustomExclusions();
  const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };

  // 除外用語をプレースホルダーに置換
  Object.keys(allExclusions).forEach(term => {
    const escapedTerm = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\b${escapedTerm}\\b`, 'gi');
    
    protectedText = protectedText.replace(regex, (match) => {
      // 数字のみのシンプルなプレースホルダー
      const placeholder = `※${placeholderIndex}※`;
      placeholders[placeholder] = match;
      placeholderIndex++;
      return placeholder;
    });
  });

  // プレースホルダーで保護されたテキストを翻訳
  let translatedText = translationFunction(protectedText);

  // プレースホルダーを元の用語に戻す
  Object.keys(placeholders).forEach(placeholder => {
    // より柔軟な復元パターン
    const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escapedPlaceholder, 'g');
    translatedText = translatedText.replace(regex, placeholders[placeholder]);
  });

  return translatedText;
}
/**
 * 保護機能付きGoogle翻訳
 * @param {string} text 翻訳する元のテキスト
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 * @returns {string} 翻訳後のテキスト
 */
function translateWithGoogle(text, src, tgt) {
  // 空文字や改行のみの場合は翻訳をスキップ
  if (!text || text.trim() === '' || text.match(/^[\n\r\s]*$/)) {
    return text;
  }

  // 除外用語を保護して翻訳
  const translatedText = protectAndTranslate(text, (protectedText) => {
    const result = LanguageApp.translate(protectedText, src, tgt);
    return result;
  });

  // 日本語への翻訳の場合、基本的な校正を適用
  let finalResult = translatedText;
  if (tgt === 'ja') {
    finalResult = basicJapaneseProofread(translatedText);
  }
  
  // API制限回避のための短いスリープ（速度向上のため50msに短縮）
  Utilities.sleep(50);
  return finalResult;
}

/**
 * 基本的な日本語校正を行う
 * @param {string} text 校正するテキスト
 * @returns {string} 校正後のテキスト
 */
function basicJapaneseProofread(text) {
  if (!text || typeof text !== 'string') return text;
  
  let corrected = text;
  
  // 1. 句読点の統一（、。を使用）
  corrected = corrected.replace(/，/g, '、');
  corrected = corrected.replace(/．/g, '。');
  
  // 2. カタカナの長音符を統一（ー）
  corrected = corrected.replace(/～/g, 'ー');
  
  // 3. 不自然な敬語の簡素化
  corrected = corrected.replace(/いたします/g, 'します');
  corrected = corrected.replace(/させていただきます/g, 'します');
  corrected = corrected.replace(/させていただく/g, 'する');
  
  // 4. 冗長な表現の簡略化
  corrected = corrected.replace(/することができます/g, 'できます');
  corrected = corrected.replace(/することが可能です/g, 'できます');
  corrected = corrected.replace(/に関して/g, 'について');
  corrected = corrected.replace(/に関しまして/g, 'について');
  
  // 5. ビジネス文書でよく使われる不自然な表現の修正
  corrected = corrected.replace(/～について説明します/g, '～を説明します');
  corrected = corrected.replace(/～に関して述べます/g, '～について述べます');
  
  // 6. 重複する助詞の修正
  corrected = corrected.replace(/のの/g, 'の');
  corrected = corrected.replace(/とと/g, 'と');
  corrected = corrected.replace(/がが/g, 'が');
  
  // 7. 不自然なスペースの除去
  corrected = corrected.replace(/([あ-んア-ンー一-龠])\s+([あ-んア-ンー一-龠])/g, '$1$2');
  
  // 8. 英語と日本語の間に適切なスペースを追加
  corrected = corrected.replace(/([a-zA-Z0-9])([あ-んア-ンー一-龠])/g, '$1 $2');
  corrected = corrected.replace(/([あ-んア-ンー一-龠])([a-zA-Z0-9])/g, '$1 $2');
  
  return corrected;
}

/**
 * 高度な日本語校正を行う（文脈を考慮した校正）
 * @param {string} text 校正するテキスト
 * @returns {string} 校正後のテキスト
 */
function advancedJapaneseProofread(text) {
  if (!text || typeof text !== 'string') return text;
  
  let corrected = basicJapaneseProofread(text);
  
  // プレゼンテーション特有の表現の改善
  const presentationReplacements = [
    // より自然なプレゼン表現に変更
    ['このスライドでは', 'ここでは'],
    ['以下に示します', '以下の通りです'],
    ['図表を示しています', '図表をご覧ください'],
    ['データを表示しています', 'データは以下の通りです'],
    ['結果を示しています', '結果は以下の通りです'],
    ['まとめると', 'まとめ'],
    ['結論として', '結論'],
    ['について考えてみましょう', 'について'],
    ['検討してみましょう', '検討します'],
  ];
  
  presentationReplacements.forEach(([from, to]) => {
    const regex = new RegExp(from, 'g');
    corrected = corrected.replace(regex, to);
  });
  
  // ビジネス用語の適切化
  const businessTermReplacements = [
    ['活用', '利用'],
    ['実施', '実行'],
    ['推進', '促進'],
    ['課題解決', '問題解決'],
    ['最適化', '改善'],
  ];
  
  businessTermReplacements.forEach(([from, to]) => {
    const regex = new RegExp(from, 'g');
    corrected = corrected.replace(regex, to);
  });
  
  return corrected;
}

/**
 * スライドのテキストを翻訳する
 * @param {Slide} slide 翻訳対象のスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSlide(slide, src, tgt) {
  translateSpeakerNotes(slide, src, tgt);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      translateShape(pageElement.asShape(), slide, src, tgt);
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      translateGroup(pageElement.asGroup(), slide, src, tgt);
    }
  }
  
  // スライド単位でのスリープ（速度向上のため250msに短縮）
  Utilities.sleep(250);
}

/**
 * シェイプのテキストを翻訳する
 * @param {Shape} shape 翻訳対象のシェイプ
 * @param {Slide} slide シェイプが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateShape(shape, slide, src, tgt) {
  /*スタイル維持のために Paragraph ごとに翻訳する*/
  for (let paragraphText of shape.getText().getParagraphs()) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    // 空文字の場合はスキップ
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const translatedText = translateWithGoogle(originalText, src, tgt);
    paragraphText.setText(translatedText);
    reduceTextSize(paragraphText);
  }
}

/**
 * シェイプのテキストを校正する
 * @param {Shape} shape 校正対象のシェイプ
 */
function proofreadShape(shape) {
  for (let paragraphText of shape.getText().getParagraphs()) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const proofreadText = advancedJapaneseProofread(originalText);
    paragraphText.setText(proofreadText);
  }
}

/**
 * グループ内のシェイプのテキストを翻訳する
 * @param {Group} group 翻訳対象のグループ
 * @param {Slide} slide グループが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateGroup(group, slide, src, tgt) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      translateShape(childElement.asShape(), slide, src, tgt);
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      translateGroup(childElement.asGroup(), slide, src, tgt);
    }
  }
}

/**
 * グループ内のシェイプのテキストを校正する
 * @param {Group} group 校正対象のグループ
 */
function proofreadGroup(group) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      proofreadShape(childElement.asShape());
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      proofreadGroup(childElement.asGroup());
    }
  }
}

/**
 * スライド内のテーブルを翻訳する
 * @param {Slide} slide 翻訳対象のスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateTableInSlide(slide, src, tgt) {
  const tables = slide.getTables();

  for (let table of tables) {
    const numRows = table.getNumRows();
    const numColumns = table.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numColumns; col++) {
        try {
          const cell = table.getCell(row, col);
          const originalText = cell.getText().asString();
          if (originalText === '' || originalText.trim() === '') {
            continue;
          }

          const translatedText = translateWithGoogle(originalText, src, tgt);
          cell.getText().setText(translatedText);
          reduceTextSize(cell.getText());
        } catch (error) {
          // 結合セルの先頭以外のセルにアクセスしようとした場合はスキップ
          console.log(`セル (${row}, ${col}) をスキップしました: ${error.message}`);
          continue;
        }
      }
    }
  }
}

/**
 * スライド内のテーブルを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function proofreadTableInSlide(slide) {
  const tables = slide.getTables();

  for (let table of tables) {
    const numRows = table.getNumRows();
    const numColumns = table.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numColumns; col++) {
        try {
          const cell = table.getCell(row, col);
          const originalText = cell.getText().asString();
          if (originalText === '' || originalText.trim() === '') {
            continue;
          }

          const proofreadText = advancedJapaneseProofread(originalText);
          cell.getText().setText(proofreadText);
        } catch (error) {
          console.log(`セル (${row}, ${col}) をスキップしました: ${error.message}`);
          continue;
        }
      }
    }
  }
}

/**
 * 全てのスライドを翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateAllSlides(src, tgt) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let slide of slides) {
    translateSlide(slide, src, tgt);
    translateTableInSlide(slide, src, tgt);
  }

  /* Google 翻訳は叩きすぎると怒られるのでスリープする（速度向上のため500msに短縮） */
  Utilities.sleep(500);
}

/**
 * 全てのスライドを校正する
 */
function proofreadAllSlides() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let slide of slides) {
    proofreadSlide(slide);
    proofreadTableInSlide(slide);
  }
}

/**
 * スライドのテキストを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function proofreadSlide(slide) {
  proofreadSpeakerNotes(slide);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      proofreadShape(pageElement.asShape());
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      proofreadGroup(pageElement.asGroup());
    }
  }
}

/**
 * 特定のスライドページを翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSpecificPage(src, tgt) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  translateSlide(slide, src, tgt);
  translateTableInSlide(slide, src, tgt);
  
  // 特定ページ翻訳後のスリープ（速度向上のため500msに短縮）
  Utilities.sleep(500);
}

/**
 * 特定のスライドページを校正する
 */
function proofreadSpecificPage() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  proofreadSlide(slide);
  proofreadTableInSlide(slide);
}

/**
 * スピーカーノートを翻訳する
 * @param {Slide} slide テキストが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSpeakerNotes(slide, src, tgt) {
  const notesPageTexts = slide.getNotesPage().getSpeakerNotesShape().getText().getParagraphs();
  for(let paragraphText of notesPageTexts) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const translatedText = translateWithGoogle(originalText, src, tgt);
    paragraphText.setText(translatedText);
  }
}

/**
 * スピーカーノートを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function proofreadSpeakerNotes(slide) {
  const notesPageTexts = slide.getNotesPage().getSpeakerNotesShape().getText().getParagraphs();
  for(let paragraphText of notesPageTexts) {
    paragraphText = paragraphText.getRange();
    const originalText = paragraphText.asString();
    
    if (!originalText || originalText.trim() === '') {
      continue;
    }
    
    const proofreadText = advancedJapaneseProofread(originalText);
    paragraphText.setText(proofreadText);
  }
}

/**
 * ユーザーからスライドのページ番号を取得する
 * @returns {?number} ページ番号 (キャンセルされた場合はnull)
 */
function getPageNumberFromUser() {
  const ui = SlidesApp.getUi();
  const response = ui.prompt(
    '翻訳こんにゃく',
    '翻訳したいスライドのページ番号を入力してください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.CANCEL) {
    console.log('キャンセルが押されたため、処理を中断しました。');
    return null;
  } else if (response.getSelectedButton() === ui.Button.CLOSE) {
    console.log('閉じるボタンが押されました。');
    return null;
  }

  const pageNumber = parseInt(response.getResponseText(), 10);
  console.log(`ページ番号は、${pageNumber} です。`);
  return pageNumber;
}

/**
 * テキストのサイズを減らす
 * @param {text} object
 */
function reduceTextSize(text) {
  if (text.asString().trim() !== "") {
    const textStyle = text.getTextStyle();
    if (textStyle.getFontSize() !== null) {
      // フォントサイズを減らす
      textStyle.setFontSize((textStyle.getFontSize() * 0.9).toFixed());
    }
  }
}
/**使用して全てのスライドを校正する*/
function databricksProofreadAllSlides() {
  if (!DATABRICKS_TOKEN || !DATABRICKS_WORKSPACE) {
    SlidesApp.getUi().alert('Databricks API設定が不完全です。「Databricks API設定」から設定してください。');
    return;
  }
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    databricksProofreadSlide(slide);
    databricksProofreadTableInSlide(slide);
    
    // API制限を考慮してスリープ
    Utilities.sleep(2000);
  }
  
  SlidesApp.getUi().alert(`${slides.length}枚のスライドの校正が完了しました。`);
}

/**
 * Databricks APIを使用して特定のスライドページを校正する
 */
function databricksProofreadSpecificPage() {
  if (!DATABRICKS_TOKEN || !DATABRICKS_WORKSPACE) {
    SlidesApp.getUi().alert('Databricks API設定が不完全です。「Databricks API設定」から設定してください。');
    return;
  }
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  databricksProofreadSlide(slide);
  databricksProofreadTableInSlide(slide);
  
  SlidesApp.getUi().alert(`スライド${pageNumber}の校正が完了しました。`);
}

/**
 * Databricks APIを使用してテキストを校正する（保護機能付き）
 * @param {string} text 校正するテキスト
 * @returns {string} 校正後のテキスト
 */
function proofreadWithDatabricks(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return text;
  }

  // カスタム除外辞書を取得
  const customExclusions = getCustomExclusions();
  const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };
  
  // 除外用語リストを作成
  const exclusionList = Object.keys(allExclusions).join(', ');

  const prompt = `以下の日本語テキストをプレゼンテーション用として自然で読みやすく校正してください。

重要な注意事項：
以下の技術用語・サービス名は絶対に変更せず、そのまま保持してください：
${exclusionList}

校正ガイドライン：
1. 不自然な敬語や冗長な表現を簡潔にする
2. プレゼンテーション向けの分かりやすい表現にする
3. 句読点を適切に使用する（、。を基本）
4. カタカナの長音符は「ー」に統一
5. 英語と日本語の間に適切なスペースを入れる
6. 文章の意味は変えずに、より自然な日本語にする
7. 上記の技術用語・サービス名は変更厳禁
8. 校正後のテキストのみを返してください

元のテキスト：${text}

校正後：`;

  const apiUrl = `https://${DATABRICKS_WORKSPACE}.cloud.databricks.com/serving-endpoints/${DATABRICKS_MODEL_ENDPOINT}/invocations`;
  
  const payload = {
    inputs: {
      messages: [
        {
          role: "system",
          content: "あなたは日本語の校正専門家です。技術用語やサービス名は絶対に変更せず、プレゼンテーション用の文書を自然で読みやすい日本語に校正してください。"
        },
        {
          role: "user",
          content: prompt
        }
      ],
      max_tokens: 1000,
      temperature: 0.3,
      top_p: 0.9
    }
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${DATABRICKS_TOKEN}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseText = response.getContentText();
    const responseJson = JSON.parse(responseText);
    
    if (responseJson.choices && responseJson.choices.length > 0) {
      let correctedText = responseJson.choices[0].message.content;
      
      // レスポンスから不要な説明文を除去
      correctedText = correctedText.replace(/^.*?校正後[：:]\s*/m, '');
      correctedText = correctedText.replace(/^.*?結果[：:]\s*/m, '');
      correctedText = correctedText.replace(/^校正された文章[：:]\s*/m, '');
      correctedText = correctedText.trim();
      
      // 改行で複数行がある場合は最初の行のみ使用（説明文除去のため）
      const lines = correctedText.split('\n');
      if (lines.length > 1 && lines[0].length > 10) {
        correctedText = lines[0];
      }
      
      return correctedText || text;
    } else if (responseJson.predictions && responseJson.predictions.length > 0) {
      // 別のレスポンス形式に対応
      let correctedText = responseJson.predictions[0];
      correctedText = correctedText.replace(/^.*?校正後[：:]\s*/m, '');
      correctedText = correctedText.trim();
      return correctedText || text;
    } else {
      console.log('Databricks API からの応答が不正です:', responseJson);
      return text;
    }
  } catch (error) {
    console.log('Databricks API エラー:', error);
    // エラーが発生した場合は基本的な校正にフォールバック
    return advancedJapaneseProofread(text);
  }
}

/**
 * Databricks APIを使用してスライドのテキストを校正する
 * @param {Slide} slide 校正対象のスライド
 */
function databricksProofreadSlide(slide) {
  databricksProofreadSpeakerNotes(slide);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      databricksProofreadShape(pageElement.asShape());
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      databricksProofreadGroup(pageElement.asGroup());
    }
  }
}
