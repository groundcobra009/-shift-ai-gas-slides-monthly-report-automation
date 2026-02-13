/**
 * スライドレポート自動生成システム (改善版)
 * Google Apps Script Implementation
 */

// ========================================
// カラーテーマ定義
// ========================================

const COLOR_THEMES = {
  green: {
    name: '緑',
    primary: '#2d8659',
    primaryDark: '#1e5d3f',
    secondary: '#4a9d73',
    accent: '#6bb88f',
    light: '#8fc9a8',
    background: '#f0f7f4',
    text: '#1a3d2a',
    chartColors: ['#2d8659', '#4a9d73', '#6bb88f', '#8fc9a8', '#a8d4bb', '#5ba67a', '#3d8b63', '#6bb88f']
  },
  monochrome: {
    name: 'モノトーン',
    primary: '#2c3e50',
    primaryDark: '#1a252f',
    secondary: '#34495e',
    accent: '#5d6d7e',
    light: '#85929e',
    background: '#ffffff',
    text: '#1a252f',
    chartColors: ['#2c3e50', '#34495e', '#5d6d7e', '#85929e', '#aeb6bf', '#d5d8dc', '#e5e7e9', '#f4f6f7']
  },
  blue: {
    name: '青',
    primary: '#2c5282',
    primaryDark: '#1e3a5f',
    secondary: '#3d6fa3',
    accent: '#5a8fc4',
    light: '#7ba8d1',
    background: '#f0f4f8',
    text: '#1e3a5f',
    chartColors: ['#2c5282', '#3d6fa3', '#5a8fc4', '#7ba8d1', '#9bbfe0', '#4a7fb8', '#2c5282', '#5a8fc4']
  },
  red: {
    name: '赤',
    primary: '#8b4a6b',
    primaryDark: '#6b3752',
    secondary: '#a86585',
    accent: '#c485a3',
    light: '#d9a5bf',
    background: '#f8f4f6',
    text: '#6b3752',
    chartColors: ['#8b4a6b', '#a86585', '#c485a3', '#d9a5bf', '#e8c4d7', '#b87595', '#8b4a6b', '#c485a3']
  }
};

function getColorTheme(themeName) {
  return COLOR_THEMES[themeName] || COLOR_THEMES.green;
}

function getAvailableColorThemes() {
  return Object.keys(COLOR_THEMES).map(key => ({
    id: key,
    name: COLOR_THEMES[key].name
  }));
}

// ========================================
// 設定管理（スクリプトプロパティ）
// ========================================

function getScriptProperties_() {
  const props = PropertiesService.getScriptProperties();
  return {
    slideTemplateId: props.getProperty('SLIDE_TEMPLATE_ID') || '',
    slideTemplateIdMonthly: props.getProperty('SLIDE_TEMPLATE_ID_MONTHLY') || '',
    slideTemplateIdYearly: props.getProperty('SLIDE_TEMPLATE_ID_YEARLY') || '',
    slideTemplateIdWeekly: props.getProperty('SLIDE_TEMPLATE_ID_WEEKLY') || '',
    currentSlideId: props.getProperty('CURRENT_SLIDE_ID') || '',
    outputFolderId: props.getProperty('OUTPUT_FOLDER_ID') || '',
    geminiApiKey: props.getProperty('GEMINI_API_KEY') || '',
    periodType: props.getProperty('PERIOD_TYPE') || 'monthly',
    colorTheme: props.getProperty('COLOR_THEME') || 'green'
  };
}

function saveScriptProperties_(config) {
  const props = PropertiesService.getScriptProperties();
  if (config.slideTemplateId !== undefined) props.setProperty('SLIDE_TEMPLATE_ID', config.slideTemplateId);
  if (config.slideTemplateIdMonthly !== undefined) props.setProperty('SLIDE_TEMPLATE_ID_MONTHLY', config.slideTemplateIdMonthly);
  if (config.slideTemplateIdYearly !== undefined) props.setProperty('SLIDE_TEMPLATE_ID_YEARLY', config.slideTemplateIdYearly);
  if (config.slideTemplateIdWeekly !== undefined) props.setProperty('SLIDE_TEMPLATE_ID_WEEKLY', config.slideTemplateIdWeekly);
  if (config.currentSlideId !== undefined) props.setProperty('CURRENT_SLIDE_ID', config.currentSlideId);
  if (config.outputFolderId !== undefined) props.setProperty('OUTPUT_FOLDER_ID', config.outputFolderId);
  if (config.geminiApiKey !== undefined) props.setProperty('GEMINI_API_KEY', config.geminiApiKey);
  if (config.periodType !== undefined) props.setProperty('PERIOD_TYPE', config.periodType);
  if (config.colorTheme !== undefined) props.setProperty('COLOR_THEME', config.colorTheme);
}

function getConfigForUI() {
  const config = getScriptProperties_();
  if (config.geminiApiKey) {
    config.geminiApiKeyMasked = maskApiKey_(config.geminiApiKey);
  } else {
    config.geminiApiKeyMasked = '';
  }
  delete config.geminiApiKey;
  config.availableThemes = getAvailableColorThemes();
  return config;
}

function saveConfigFromUI(config) {
  try {
    if (config.geminiApiKey && !config.geminiApiKey.includes('*')) {
      saveScriptProperties_({ geminiApiKey: config.geminiApiKey });
    }
    saveScriptProperties_({
      slideTemplateId: config.slideTemplateId,
      outputFolderId: config.outputFolderId
    });
    return { success: true, message: '設定を保存しました' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function maskApiKey_(apiKey) {
  if (!apiKey || apiKey.length < 8) return '********';
  return apiKey.substring(0, 4) + '****************' + apiKey.substring(apiKey.length - 4);
}

// ========================================
// 初期セットアップ
// ========================================

function setupInitialEnvironment() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolder = ssFile.getParents().next();

    const results = {
      templateDeleted: false,
      folderDeleted: false,
      templateCreated: false,
      folderCreated: false,
      templateId: '',
      folderId: '',
      templateUrl: '',
      folderUrl: ''
    };

    const config = getScriptProperties_();

    // ========================================
    // STEP 1: 既存のリソースを完全削除
    // ========================================

    // 1-1. スクリプトプロパティに保存されているテンプレートを削除
    if (config.slideTemplateId) {
      try {
        const oldTemplate = DriveApp.getFileById(config.slideTemplateId);
        oldTemplate.setTrashed(true);
        results.templateDeleted = true;
      } catch (e) {
        // 既に削除済みの場合は無視
      }
    }

    // 1-2. 同一階層にある「📊 レポートテンプレート」という名前のファイルを全て削除
    const existingTemplates = parentFolder.getFilesByName('📊 レポートテンプレート');
    while (existingTemplates.hasNext()) {
      const file = existingTemplates.next();
      file.setTrashed(true);
      results.templateDeleted = true;
    }

    // 1-3. スクリプトプロパティに保存されている出力フォルダを削除
    if (config.outputFolderId) {
      try {
        const oldFolder = DriveApp.getFolderById(config.outputFolderId);
        oldFolder.setTrashed(true);
        results.folderDeleted = true;
      } catch (e) {
        // 既に削除済みの場合は無視
      }
    }

    // 1-4. 同一階層にある「📁 レポート出力」という名前のフォルダを全て削除
    const existingFolders = parentFolder.getFoldersByName('📁 レポート出力');
    while (existingFolders.hasNext()) {
      const folder = existingFolders.next();
      folder.setTrashed(true);
      results.folderDeleted = true;
    }

    // 1-5. スクリプトプロパティをクリア
    saveScriptProperties_({
      slideTemplateId: '',
      outputFolderId: '',
      currentSlideId: ''
    });

    // ========================================
    // STEP 2: 新しいリソースを作成
    // ========================================

    // 2-1. 新しいテンプレートを作成（現在のカラーテーマを使用）
    const themeName = config.colorTheme || 'green';
    
    // 月次、年次、週次のテンプレートを作成
    const monthlyTemplate = createSlideTemplate_(parentFolder, themeName, 'monthly');
    const yearlyTemplate = createSlideTemplate_(parentFolder, themeName, 'yearly');
    const weeklyTemplate = createSlideTemplate_(parentFolder, themeName, 'weekly');
    
    results.templateCreated = true;
    results.templateId = monthlyTemplate.getId();
    results.templateUrl = monthlyTemplate.getUrl();
    
    // 各テンプレートIDを保存
    saveScriptProperties_({ 
      slideTemplateId: monthlyTemplate.getId(),
      slideTemplateIdMonthly: monthlyTemplate.getId(),
      slideTemplateIdYearly: yearlyTemplate.getId(),
      slideTemplateIdWeekly: weeklyTemplate.getId()
    });

    // 2-2. 新しい出力フォルダを作成
    const outputFolder = createOutputFolder_(parentFolder);
    results.folderCreated = true;
    results.folderId = outputFolder.getId();
    results.folderUrl = outputFolder.getUrl();
    saveScriptProperties_({ outputFolderId: outputFolder.getId() });

    return {
      success: true,
      message: (results.templateDeleted || results.folderDeleted) ?
        '既存のリソースを削除し、新しい環境をセットアップしました' :
        '初期セットアップが完了しました',
      results: results
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

function createSlideTemplate_(folder, themeName = 'green', periodType = 'monthly') {
  const theme = getColorTheme(themeName);
  
  // periodTypeに応じてテンプレート名を変更
  let templateName = '📊 レポートテンプレート';
  if (periodType === 'yearly') {
    templateName = '📊 年次レポートテンプレート';
  } else if (periodType === 'weekly') {
    templateName = '📊 週次レポートテンプレート';
  } else {
    templateName = '📊 月次レポートテンプレート';
  }
  
  const presentation = SlidesApp.create(templateName);
  const slides = presentation.getSlides();

  // Slide 1: 表紙（モダンデザイン）
  const slide1 = slides[0];

  // デフォルトのシェイプを安全に削除
  try {
    const shapes = slide1.getShapes();
    for (let i = shapes.length - 1; i >= 0; i--) {
      try {
        shapes[i].remove();
      } catch (e) {
        // シェイプ削除に失敗しても続行
      }
    }
  } catch (e) {
    // シェイプ取得に失敗しても続行
  }

  // 背景グラデーション風の装飾（0.1を追加して確実に正の値にする）
  const bgShape1 = slide1.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 540);
  bgShape1.getFill().setSolidFill(theme.primary);
  bgShape1.getBorder().setTransparent();

  const bgShape2 = slide1.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 300, 720, 240);
  bgShape2.getFill().setSolidFill(theme.primaryDark);
  bgShape2.getBorder().setTransparent();

  // アクセント円
  const circle1 = slide1.insertShape(SlidesApp.ShapeType.ELLIPSE, -100, -100, 300, 300);
  circle1.getFill().setSolidFill(theme.accent);
  circle1.getBorder().setTransparent();

  const circle2 = slide1.insertShape(SlidesApp.ShapeType.ELLIPSE, 520, 340, 300, 300);
  circle2.getFill().setSolidFill(theme.secondary);
  circle2.getBorder().setTransparent();

  // タイトル（白文字）
  const titleBox1 = slide1.insertTextBox('{{reportTitle}}', 60, 180, 600, 100);
  titleBox1.getText().getTextStyle()
    .setFontSize(56)
    .setBold(true)
    .setForegroundColor('#ffffff')
    .setFontFamily('Arial');

  // サブタイトル（白文字）
  const subtitleBox1 = slide1.insertTextBox('{{period}}', 60, 280, 600, 70);
  subtitleBox1.getText().getTextStyle()
    .setFontSize(36)
    .setForegroundColor('#ffffff')
    .setFontFamily('Arial');

  // タイムスタンプ（右下・白文字）
  const timestampBox1 = slide1.insertTextBox('Generated at {{generatedAt}}', 450, 490, 250, 30);
  timestampBox1.getText().getTextStyle()
    .setFontSize(12)
    .setForegroundColor('#ffffff');

  // Slide 2: サマリー（カード風デザイン）
  const slide2 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // 背景
  const bgSlide2 = slide2.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 540);
  bgSlide2.getFill().setSolidFill(theme.background);
  bgSlide2.getBorder().setTransparent();

  // ヘッダー帯
  const headerBand = slide2.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 80);
  headerBand.getFill().setSolidFill(theme.primary);
  headerBand.getBorder().setTransparent();

  const titleBox2 = slide2.insertTextBox('📊 売上サマリー', 40, 20, 640, 50);
  titleBox2.getText().getTextStyle()
    .setFontSize(32)
    .setBold(true)
    .setForegroundColor('#ffffff')
    .setFontFamily('Arial');

  // カード背景
  const cardBg = slide2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 100, 640, 410);
  cardBg.getFill().setSolidFill('#ffffff');
  cardBg.getBorder().setTransparent();

  const summaryText = `💰 合計売上: {{totalSales}}
📈 {{growthRateLabel}}: {{totalSalesChange}}

🏆 トップ地域: {{topRegion}} ({{topRegionSales}})
👤 トップ担当者: {{topPerson}} ({{topPersonSales}})

💡 考察:
{{aiComment}}`;

  const summaryBox = slide2.insertTextBox(summaryText, 70, 130, 580, 350);
  summaryBox.getText().getTextStyle()
    .setFontSize(20)
    .setForegroundColor(theme.text)
    .setFontFamily('Arial');

  // Slide 3: 地域別売上グラフ
  const slide3 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  const bgSlide3 = slide3.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 540);
  bgSlide3.getFill().setSolidFill(theme.background);
  bgSlide3.getBorder().setTransparent();

  const headerBand3 = slide3.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 80);
  headerBand3.getFill().setSolidFill(theme.primaryDark);
  headerBand3.getBorder().setTransparent();

  const titleBox3 = slide3.insertTextBox('🌏 地域別売上', 40, 20, 640, 50);
  titleBox3.getText().getTextStyle()
    .setFontSize(32)
    .setBold(true)
    .setForegroundColor('#ffffff')
    .setFontFamily('Arial');

  // Slide 4: 担当者別売上グラフ
  const slide4 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  const bgSlide4 = slide4.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 540);
  bgSlide4.getFill().setSolidFill(theme.background);
  bgSlide4.getBorder().setTransparent();

  const headerBand4 = slide4.insertShape(SlidesApp.ShapeType.RECTANGLE, 0.1, 0.1, 720, 80);
  headerBand4.getFill().setSolidFill(theme.primaryDark);
  headerBand4.getBorder().setTransparent();

  const titleBox4 = slide4.insertTextBox('👥 担当者別売上', 40, 20, 640, 50);
  titleBox4.getText().getTextStyle()
    .setFontSize(32)
    .setBold(true)
    .setForegroundColor('#ffffff')
    .setFontFamily('Arial');

  // ファイルを移動
  const file = DriveApp.getFileById(presentation.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return file;
}

/**
 * スライドテンプレートのみを作成
 */
function createSlideTemplateOnly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolder = ssFile.getParents().next();

    const config = getScriptProperties_();

    // STEP 1: 既存テンプレートの確認と削除
    let templateExists = false;

    // 1-1. スクリプトプロパティに保存されているテンプレートIDを確認
    if (config.slideTemplateId) {
      try {
        const oldTemplate = DriveApp.getFileById(config.slideTemplateId);
        // テンプレートが存在する場合のみ削除
        oldTemplate.setTrashed(true);
        templateExists = true;
      } catch (e) {
        // ファイルが見つからない場合は無視
      }
    }

    // 1-2. 同一階層にある「📊 レポートテンプレート」という名前のファイルを確認
    const existingTemplates = parentFolder.getFilesByName('📊 レポートテンプレート');
    while (existingTemplates.hasNext()) {
      const file = existingTemplates.next();
      file.setTrashed(true);
      templateExists = true;
    }

    // 1-3. スクリプトプロパティをクリア
    saveScriptProperties_({ slideTemplateId: '' });

    // STEP 2: 新しいテンプレートを作成（現在のカラーテーマを使用）
    const themeName = config.colorTheme || 'green';
    
    // 月次テンプレートを作成（デフォルト）
    const template = createSlideTemplate_(parentFolder, themeName, 'monthly');
    saveScriptProperties_({ 
      slideTemplateId: template.getId(),
      slideTemplateIdMonthly: template.getId()
    });

    return {
      success: true,
      message: templateExists ?
        '既存のテンプレートを削除し、新しいテンプレートを作成しました' :
        'スライドテンプレートを作成しました',
      templateId: template.getId(),
      templateUrl: template.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: `テンプレート作成エラー: ${error.toString()}`
    };
  }
}

/**
 * 出力フォルダのみを作成
 */
function createOutputFolderOnly() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolder = ssFile.getParents().next();

    const config = getScriptProperties_();

    // 既存の出力フォルダを削除（存在する場合のみ）
    if (config.outputFolderId) {
      try {
        const oldFolder = DriveApp.getFolderById(config.outputFolderId);
        oldFolder.setTrashed(true);
      } catch (e) {
        // 既に削除済みの場合は無視
      }
    }

    // 同一階層の既存フォルダも削除
    const existingFolders = parentFolder.getFoldersByName('📁 レポート出力');
    while (existingFolders.hasNext()) {
      const folder = existingFolders.next();
      folder.setTrashed(true);
    }

    // 新しいフォルダを作成
    const outputFolder = createOutputFolder_(parentFolder);
    saveScriptProperties_({ outputFolderId: outputFolder.getId() });

    return {
      success: true,
      message: '出力フォルダを作成しました',
      folderId: outputFolder.getId(),
      folderUrl: outputFolder.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * 既存のテンプレートを更新（再作成）
 * テンプレートIDが未設定の場合は新規作成
 */
function updateSlideTemplate() {
  try {
    const config = getScriptProperties_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolder = ssFile.getParents().next();

    let isNewCreation = false;

    // 既存のテンプレートを削除（存在する場合のみ）
    if (config.slideTemplateId) {
      try {
        const oldTemplate = DriveApp.getFileById(config.slideTemplateId);
        oldTemplate.setTrashed(true);
      } catch (e) {
        // 既に削除済みの場合は無視
      }
    } else {
      // テンプレートIDが空の場合は新規作成
      isNewCreation = true;
    }

    // スクリプトプロパティから古いIDを明示的にクリア
    saveScriptProperties_({ slideTemplateId: '' });

    // 新しいテンプレートを作成（現在のカラーテーマを使用）
    const themeName = config.colorTheme || 'green';
    
    // 月次、年次、週次のテンプレートを作成
    const monthlyTemplate = createSlideTemplate_(parentFolder, themeName, 'monthly');
    const yearlyTemplate = createSlideTemplate_(parentFolder, themeName, 'yearly');
    const weeklyTemplate = createSlideTemplate_(parentFolder, themeName, 'weekly');

    // 新しいテンプレートIDをスクリプトプロパティに保存（紐付け更新）
    saveScriptProperties_({ 
      slideTemplateId: monthlyTemplate.getId(),
      slideTemplateIdMonthly: monthlyTemplate.getId(),
      slideTemplateIdYearly: yearlyTemplate.getId(),
      slideTemplateIdWeekly: weeklyTemplate.getId()
    });

    return {
      success: true,
      message: isNewCreation ?
        'スライドテンプレートを新規作成しました' :
        'スライドテンプレートを更新しました',
      templateId: newTemplate.getId(),
      templateUrl: newTemplate.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

function createOutputFolder_(parentFolder) {
  return parentFolder.createFolder('📁 レポート出力');
}

/**
 * 既存テンプレートのカラーテーマを変更
 */
function changeTemplateColorTheme(themeName) {
  try {
    const config = getScriptProperties_();
    
    if (!config.slideTemplateId) {
      return {
        success: false,
        message: 'テンプレートが作成されていません。まずテンプレートを作成してください。'
      };
    }

    // テーマが有効か確認
    if (!COLOR_THEMES[themeName]) {
      return {
        success: false,
        message: `無効なテーマ名です: ${themeName}`
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parentFolder = ssFile.getParents().next();

    // 既存テンプレートを削除（月次/年次/週次すべて）
    const templateIds = [
      config.slideTemplateId,
      config.slideTemplateIdMonthly,
      config.slideTemplateIdYearly,
      config.slideTemplateIdWeekly
    ].filter(id => id);
    
    templateIds.forEach(templateId => {
      try {
        const oldTemplate = DriveApp.getFileById(templateId);
        oldTemplate.setTrashed(true);
      } catch (e) {
        // 既に削除済みの場合は無視
      }
    });

    // 新しいテーマでテンプレートを作成（月次/年次/週次すべて）
    const monthlyTemplate = createSlideTemplate_(parentFolder, themeName, 'monthly');
    const yearlyTemplate = createSlideTemplate_(parentFolder, themeName, 'yearly');
    const weeklyTemplate = createSlideTemplate_(parentFolder, themeName, 'weekly');
    
    // カラーテーマとテンプレートIDを保存
    saveScriptProperties_({ 
      slideTemplateId: monthlyTemplate.getId(),
      slideTemplateIdMonthly: monthlyTemplate.getId(),
      slideTemplateIdYearly: yearlyTemplate.getId(),
      slideTemplateIdWeekly: weeklyTemplate.getId(),
      colorTheme: themeName
    });

    // 集計シートが存在する場合は、グラフの色も更新
    const rawSheet = ss.getSheetByName('RawSalesData');
    if (rawSheet && rawSheet.getLastRow() > 1) {
      try {
        refreshAggregationSheets();
      } catch (e) {
        // 集計シートの更新に失敗しても続行（テンプレートは作成済み）
        Logger.log('集計シートの更新エラー: ' + e.toString());
      }
    }

    return {
      success: true,
      message: `カラーテーマを「${COLOR_THEMES[themeName].name}」に変更しました。月次/年次/週次のテンプレートとグラフの色も更新されました。`,
      templateId: monthlyTemplate.getId(),
      templateUrl: monthlyTemplate.getUrl(),
      themeName: themeName
    };
  } catch (error) {
    return {
      success: false,
      message: `カラーテーマ変更エラー: ${error.toString()}`
    };
  }
}

/**
 * テンプレートの{{period}}プレースホルダーを削除する
 */
function fixTemplatesRemovePeriodPlaceholder() {
  try {
    const config = getScriptProperties_();
    const templateIds = [
      { id: config.slideTemplateIdMonthly, name: '月次' },
      { id: config.slideTemplateIdYearly, name: '年次' },
      { id: config.slideTemplateIdWeekly, name: '週次' },
      { id: config.slideTemplateId, name: 'デフォルト' }
    ];

    let updatedCount = 0;
    const results = [];

    for (const template of templateIds) {
      if (!template.id) continue;

      try {
        const presentation = SlidesApp.openById(template.id);
        const slides = presentation.getSlides();

        if (slides.length > 0) {
          const slide1 = slides[0];
          const shapes = slide1.getShapes();
          let found = false;

          for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
              const text = shape.getText().asString();

              // {{period}}プレースホルダーを削除
              if (text.includes('{{period}}')) {
                shape.remove();
                found = true;
              }
            }
          }

          if (found) {
            updatedCount++;
            results.push(`${template.name}: {{period}}を削除しました`);
          } else {
            results.push(`${template.name}: {{period}}が見つかりませんでした`);
          }
        }
      } catch (e) {
        results.push(`${template.name}: エラー - ${e.toString()}`);
      }
    }

    return {
      success: true,
      message: `テンプレート修正完了（${updatedCount}件更新）\n\n${results.join('\n')}`
    };
  } catch (error) {
    return {
      success: false,
      message: `テンプレート修正エラー: ${error.toString()}`
    };
  }
}

/**
 * すべてのテンプレートを削除する
 */
function deleteAllTemplates() {
  try {
    const config = getScriptProperties_();
    const templateIds = [
      { id: config.slideTemplateIdMonthly, name: '月次' },
      { id: config.slideTemplateIdYearly, name: '年次' },
      { id: config.slideTemplateIdWeekly, name: '週次' },
      { id: config.slideTemplateId, name: 'デフォルト' }
    ];

    let deletedCount = 0;
    const results = [];

    for (const template of templateIds) {
      if (!template.id) {
        results.push(`${template.name}: テンプレートIDが設定されていません`);
        continue;
      }

      try {
        const file = DriveApp.getFileById(template.id);
        file.setTrashed(true);
        deletedCount++;
        results.push(`${template.name}: 削除しました`);
      } catch (e) {
        // ファイルが見つからない場合は既に削除されている
        if (e.toString().includes('File not found')) {
          results.push(`${template.name}: 既に削除されています`);
        } else {
          results.push(`${template.name}: エラー - ${e.toString()}`);
        }
      }
    }

    // スクリプトプロパティをクリア
    saveScriptProperties_({
      slideTemplateId: '',
      slideTemplateIdMonthly: '',
      slideTemplateIdYearly: '',
      slideTemplateIdWeekly: ''
    });

    return {
      success: true,
      message: `テンプレート削除完了（${deletedCount}件削除）\n\n${results.join('\n')}`
    };
  } catch (error) {
    return {
      success: false,
      message: `テンプレート削除エラー: ${error.toString()}`
    };
  }
}

/**
 * テンプレートの表紙タイトルを{{reportTitle}}プレースホルダーに統一する
 */
function fixTemplatesReportTitlePlaceholder() {
  try {
    const config = getScriptProperties_();
    const templateIds = [
      { id: config.slideTemplateIdMonthly, name: '月次' },
      { id: config.slideTemplateIdYearly, name: '年次' },
      { id: config.slideTemplateIdWeekly, name: '週次' },
      { id: config.slideTemplateId, name: 'デフォルト' }
    ];

    let updatedCount = 0;
    const results = [];

    for (const template of templateIds) {
      if (!template.id) continue;

      try {
        const presentation = SlidesApp.openById(template.id);
        const slides = presentation.getSlides();

        if (slides.length > 0) {
          const slide1 = slides[0];
          const shapes = slide1.getShapes();
          let titleUpdated = false;
          let largestTextBox = null;
          let largestFontSize = 0;

          // すべてのテキストボックスをチェック
          for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
              const text = shape.getText().asString();
              
              // 既に{{reportTitle}}プレースホルダーがある場合はスキップ
              if (text.includes('{{reportTitle}}')) {
                titleUpdated = true;
                results.push(`${template.name}: 既に{{reportTitle}}が設定されています`);
                break;
              }
              
              // 最大フォントサイズのテキストボックスを記録
              try {
                const fontSize = shape.getText().getTextStyle().getFontSize();
                if (fontSize > largestFontSize && text.trim() !== '') {
                  largestFontSize = fontSize;
                  largestTextBox = shape;
                }
              } catch (e) {
                // フォントサイズ取得エラーは無視
              }
              
              // 「月次」「年次」「週次」「レポート」を含むテキストボックスをタイトルとして更新
              if ((text.includes('月次') || text.includes('年次') || text.includes('週次')) && text.includes('レポート')) {
                shape.getText().setText('{{reportTitle}}');
                titleUpdated = true;
                updatedCount++;
                results.push(`${template.name}: タイトルを{{reportTitle}}に更新しました`);
                break;
              }
            }
          }

          // タイトルが見つからなかった場合、最大フォントサイズのテキストボックスを更新
          if (!titleUpdated && largestTextBox && largestFontSize >= 40) {
            largestTextBox.getText().setText('{{reportTitle}}');
            titleUpdated = true;
            updatedCount++;
            results.push(`${template.name}: 最大フォントサイズのテキストボックスを{{reportTitle}}に更新しました`);
          }
          
          if (!titleUpdated) {
            results.push(`${template.name}: タイトルテキストボックスが見つかりませんでした`);
          }
        }
      } catch (e) {
        results.push(`${template.name}: エラー - ${e.toString()}`);
      }
    }

    return {
      success: true,
      message: `テンプレートタイトル修正完了（${updatedCount}件更新）\n\n${results.join('\n')}`
    };
  } catch (error) {
    return {
      success: false,
      message: `テンプレートタイトル修正エラー: ${error.toString()}`
    };
  }
}

// ========================================
// エントリポイント
// ========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 スライドレポート')
    .addItem('⚙️ 設定', 'showSettingsSidebar')
    .addSeparator()
    .addItem('🎲 ダミーデータ生成', 'showDummyDataDialog')
    .addItem('📊 レポート生成', 'showReportDialog')
    .addSeparator()
    .addItem('❓ ヘルプ', 'showHelpDialog')
    .addToUi();
}

function showSettingsSidebar() {
  try {
    Logger.log('showSettingsSidebar: 開始');
    const html = HtmlService.createHtmlOutputFromFile('ui/SettingsSidebar')
      .setTitle('⚙️ 設定')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log('showSettingsSidebar: 完了');
  } catch (e) {
    Logger.log('showSettingsSidebar エラー: ' + e.message + '\nStack: ' + e.stack);
    SpreadsheetApp.getUi().alert('設定サイドバーの表示に失敗しました: ' + e.message);
  }
}

function showDummyDataDialog() {
  try {
    Logger.log('showDummyDataDialog: 開始');
    const html = HtmlService.createHtmlOutputFromFile('ui/dialogs/DummyDataDialog')
      .setWidth(700)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, '🎲 リアルなダミーデータ生成');
    Logger.log('showDummyDataDialog: 完了');
  } catch (e) {
    Logger.log('showDummyDataDialog エラー: ' + e.message + '\nStack: ' + e.stack);
    SpreadsheetApp.getUi().alert('ダミーデータダイアログの表示に失敗しました: ' + e.message);
  }
}

function showReportDialog() {
  try {
    Logger.log('showReportDialog: 開始');
    const html = HtmlService.createHtmlOutputFromFile('ui/MainSidebar')
      .setWidth(650)
      .setHeight(700);
    SpreadsheetApp.getUi().showModalDialog(html, '📊 レポート生成');
    Logger.log('showReportDialog: 完了');
  } catch (e) {
    Logger.log('showReportDialog エラー: ' + e.message + '\nStack: ' + e.stack);
    SpreadsheetApp.getUi().alert('レポートダイアログの表示に失敗しました: ' + e.message);
  }
}

function showHelpDialog() {
  try {
    Logger.log('showHelpDialog: 開始');
    const html = HtmlService.createHtmlOutputFromFile('ui/dialogs/HelpDialog')
      .setWidth(850)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, '📚 使い方ガイド');
    Logger.log('showHelpDialog: 完了');
  } catch (e) {
    Logger.log('showHelpDialog エラー: ' + e.message + '\nStack: ' + e.stack);
    SpreadsheetApp.getUi().alert('ヘルプダイアログの表示に失敗しました: ' + e.message);
  }
}

// ========================================
// CSVインポート＆集計機能
// ========================================

/**
 * CSVデータをインポートして集計
 */
function importAndAggregateSalesData(csvText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 生データシートを作成
    let rawSheet = ss.getSheetByName('RawSalesData');
    if (rawSheet) {
      ss.deleteSheet(rawSheet);
    }
    rawSheet = ss.insertSheet('RawSalesData');

    // CSVをパース
    const rows = Utilities.parseCsv(csvText);
    if (rows.length === 0) {
      throw new Error('CSVデータが空です');
    }

    // データを書き込み
    rawSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    rawSheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');

    // 集計シートを数式ベースで作成
    createAggregationSheets_();

    // フィルタを追加（ユーザーがスプレッドシート上で期間を選択可能に）
    addFilterViewToRawData_();

    return {
      success: true,
      message: `${rows.length - 1}件のデータをインポートし、集計シートを作成しました。\nRawSalesDataシートでフィルタを使って期間を絞り込めます。`,
      records: rows.length - 1
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * グラフを全て削除
 */
function removeAllCharts_(sheet) {
  const charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));
}

/**
 * 集計シートを数式ベースで作成（動的対応）
 */
function createAggregationSheets_() {
  createRegionSheet_();
  createPersonSheet_();
  createProductSheet_();
  createCategorySheet_();
  createMonthlySheet_();
}

/**
 * シート更新専用関数（RawSalesData更新後に実行）
 */
function refreshAggregationSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('RawSalesData');

    if (!rawSheet) {
      return {
        success: false,
        message: 'RawSalesDataシートが見つかりません'
      };
    }

    // 既存の集計シートを削除して再作成
    ['RegionalSales', 'PersonSales', 'ProductSales', 'CategorySales', 'MonthlySales'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });

    createAggregationSheets_();

    // フィルタビューを追加（スプレッドシートの標準機能）
    addFilterViewToRawData_();

    return {
      success: true,
      message: '集計シートを更新しました'
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * RawSalesDataシートにフィルタビューを追加
 */
function addFilterViewToRawData_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('RawSalesData');

    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      return;
    }

    // 既存のフィルタを削除
    const existingFilter = rawSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    // データ範囲全体にフィルタを作成
    const lastRow = rawSheet.getLastRow();
    const lastColumn = rawSheet.getLastColumn();
    const range = rawSheet.getRange(1, 1, lastRow, lastColumn);

    // 標準フィルタを作成（ユーザーがUI上で操作可能）
    const filter = range.createFilter();

    Logger.log('RawSalesDataにフィルタを追加しました');
  } catch (error) {
    Logger.log('フィルタ追加エラー: ' + error.toString());
  }
}

function createRegionSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('RegionalSales');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('RegionalSales');

  // 現在のカラーテーマを取得
  const config = getScriptProperties_();
  const theme = getColorTheme(config.colorTheme || 'green');

  // QUERY関数で集計（ヘッダーも含めて自動生成）
  const formula = '=QUERY(RawSalesData!A:H, "SELECT B, SUM(H) WHERE B IS NOT NULL GROUP BY B ORDER BY SUM(H) DESC LABEL B \'地域\', SUM(H) \'売上\'", 1)';
  sheet.getRange('A1').setFormula(formula);

  // スタイル設定
  Utilities.sleep(1500); // 数式の計算を待つ
  
  // C列に万円単位の売上を計算（チャート用）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange('C1').setValue('売上(万円)');
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(`C${i}`).setFormula(`=B${i}/10000`);
    }
  }

  sheet.getRange('A1:C1').setFontWeight('bold').setBackground(theme.primary).setFontColor('#ffffff');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);

  // 数値フォーマット（B列が売上、C列が万円単位）
  sheet.getRange('B2:B').setNumberFormat('#,##0');
  sheet.getRange('C2:C').setNumberFormat('#,##0');

  // チャート作成（既存のグラフを削除してから作成）
  removeAllCharts_(sheet);

  if (lastRow > 1) {
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange('A2:A' + lastRow))  // 地域名
      .addRange(sheet.getRange('C2:C' + lastRow))  // 万円単位の売上
      .setPosition(2, 4, 0, 0)
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('vAxis', {
        title: '売上額 (万円)',
        format: '#,##0'  // 万円単位で整数表示
      })
      .setOption('colors', theme.chartColors);

    sheet.insertChart(chartBuilder.build());
  }
}

function createPersonSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PersonSales');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('PersonSales');

  // 現在のカラーテーマを取得
  const config = getScriptProperties_();
  const theme = getColorTheme(config.colorTheme || 'green');

  // QUERY関数で集計（ヘッダーも含めて自動生成）
  const formula = '=QUERY(RawSalesData!A:H, "SELECT C, SUM(H), COUNT(H) WHERE C IS NOT NULL GROUP BY C ORDER BY SUM(H) DESC LABEL C \'担当者\', SUM(H) \'売上\', COUNT(H) \'件数\'", 1)';
  sheet.getRange('A1').setFormula(formula);

  // スタイル設定
  Utilities.sleep(1500); // 数式の計算を待つ

  // E列に万円単位の売上を計算（チャート用）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange('E1').setValue('売上(万円)');
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(`E${i}`).setFormula(`=B${i}/10000`);
    }
  }

  // D1に平均単価ヘッダーを追加
  sheet.getRange('D1').setValue('平均単価');

  // D列に平均単価の数式（B列 / C列）
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(`D${i}`).setFormula(`=IF(C${i}>0, B${i}/C${i}, 0)`);
    }
  }

  sheet.getRange('A1:E1').setFontWeight('bold').setBackground(theme.primary).setFontColor('#ffffff');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);

  // 数値フォーマット
  sheet.getRange('B2:B').setNumberFormat('#,##0');
  sheet.getRange('D2:D').setNumberFormat('#,##0');
  sheet.getRange('E2:E').setNumberFormat('#,##0');

  // チャート作成
  removeAllCharts_(sheet);

  if (lastRow > 1) {
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange('A2:A' + lastRow))  // 担当者名
      .addRange(sheet.getRange('E2:E' + lastRow))  // 万円単位の売上
      .setPosition(2, 6, 0, 0)
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('hAxis', {
        title: '売上額 (万円)',
        format: '#,##0'  // 万円単位で整数表示
      })
      .setOption('vAxis', { title: '担当者' })
      .setOption('colors', theme.chartColors);

    sheet.insertChart(chartBuilder.build());
  }
}

function createProductSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ProductSales');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('ProductSales');

  // 現在のカラーテーマを取得
  const config = getScriptProperties_();
  const theme = getColorTheme(config.colorTheme || 'green');

  // QUERY関数で集計（ヘッダーも含めて自動生成）
  const formula = '=QUERY(RawSalesData!A:H, "SELECT D, SUM(H), SUM(F) WHERE D IS NOT NULL GROUP BY D ORDER BY SUM(H) DESC LABEL D \'製品\', SUM(H) \'売上\', SUM(F) \'販売数\'", 1)';
  sheet.getRange('A1').setFormula(formula);

  // スタイル設定
  Utilities.sleep(1500); // 数式の計算を待つ

  // D1に平均単価ヘッダーを追加
  sheet.getRange('D1').setValue('平均単価');

  // D列に平均単価の数式
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(`D${i}`).setFormula(`=IF(C${i}>0, B${i}/C${i}, 0)`);
    }
  }

  sheet.getRange('A1:D1').setFontWeight('bold').setBackground(theme.primary).setFontColor('#ffffff');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 120);

  // 数値フォーマット
  sheet.getRange('B2:B').setNumberFormat('#,##0');
  sheet.getRange('D2:D').setNumberFormat('#,##0');

  // チャート作成
  removeAllCharts_(sheet);

  const chartLastRow = sheet.getLastRow();
  if (chartLastRow > 1) {
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('A1:B' + chartLastRow))
      .setPosition(2, 6, 0, 0)
      .setOption('title', '商品別売上構成')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('colors', theme.chartColors);

    sheet.insertChart(chartBuilder.build());
  }
}

function createCategorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('CategorySales');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('CategorySales');

  // 現在のカラーテーマを取得
  const config = getScriptProperties_();
  const theme = getColorTheme(config.colorTheme || 'green');

  // QUERY関数で集計（ヘッダーも含めて自動生成）
  const formula = '=QUERY(RawSalesData!A:H, "SELECT E, SUM(H) WHERE E IS NOT NULL GROUP BY E ORDER BY SUM(H) DESC LABEL E \'カテゴリ\', SUM(H) \'売上\'", 1)';
  sheet.getRange('A1').setFormula(formula);

  // スタイル設定
  Utilities.sleep(1500); // 数式の計算を待つ
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground(theme.primary).setFontColor('#ffffff');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 130);

  // 数値フォーマット
  sheet.getRange('B2:B').setNumberFormat('#,##0');

  // チャート作成
  removeAllCharts_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange('A1:B' + lastRow))
      .setPosition(2, 4, 0, 0)
      .setOption('title', 'カテゴリ別売上構成')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('colors', theme.chartColors);

    sheet.insertChart(chartBuilder.build());
  }
}

function createMonthlySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('MonthlySales');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('MonthlySales');

  // 現在のカラーテーマを取得
  const config = getScriptProperties_();
  const theme = getColorTheme(config.colorTheme || 'green');

  // QUERY関数で年月ごとの売上を集計（YEAR, MONTHを別々に取得）
  const formula = '=QUERY(RawSalesData!A:H, "SELECT YEAR(A), MONTH(A), SUM(H) WHERE A IS NOT NULL GROUP BY YEAR(A), MONTH(A) ORDER BY YEAR(A), MONTH(A) LABEL YEAR(A) \'年\', MONTH(A) \'月\', SUM(H) \'売上\'", 1)';
  sheet.getRange('B1').setFormula(formula);

  // スタイル設定
  Utilities.sleep(1500); // 数式の計算を待つ

  // A1に年月ヘッダーを追加
  sheet.getRange('A1').setValue('年月');

  const lastRow = sheet.getLastRow();

  // A列に年月を結合（例: 2025-01）
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      sheet.getRange(`A${i}`).setFormula(`=B${i}&"-"&TEXT(C${i},"00")`);
    }
  }

  // E1, F1, G1にヘッダーを追加
  sheet.getRange('E1').setValue('前月比');
  sheet.getRange('F1').setValue('前月比率');
  sheet.getRange('G1').setValue('前年同月比');

  // E列: 前月比（差額）、F列: 前月比率、G列: 前年同月比率
  // 新しい列構成: A=年月, B=年, C=月, D=売上, E=前月比, F=前月比率, G=前年同月比
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      // 前月比（D列が売上）
      if (i === 2) {
        sheet.getRange(`E${i}`).setValue('-');
        sheet.getRange(`F${i}`).setValue('-');
      } else {
        sheet.getRange(`E${i}`).setFormula(`=IF(D${i}>0, D${i}-D${i-1}, "")`);
        sheet.getRange(`F${i}`).setFormula(`=IF(D${i-1}>0, (D${i}/D${i-1}-1), "")`);
      }

      // 前年同月比（12ヶ月前）
      if (i > 13) {
        sheet.getRange(`G${i}`).setFormula(`=IF(D${i-12}>0, (D${i}/D${i-12}-1), "")`);
      } else {
        sheet.getRange(`G${i}`).setValue('');
      }
    }
  }

  sheet.getRange('A1:G1').setFontWeight('bold').setBackground(theme.primary).setFontColor('#ffffff');
  sheet.setColumnWidth(1, 100);  // 年月
  sheet.setColumnWidth(2, 60);   // 年
  sheet.setColumnWidth(3, 60);   // 月
  sheet.setColumnWidth(4, 120);  // 売上
  sheet.setColumnWidth(5, 100);  // 前月比
  sheet.setColumnWidth(6, 100);  // 前月比率
  sheet.setColumnWidth(7, 100);  // 前年同月比

  // 数値フォーマット
  sheet.getRange('D2:D').setNumberFormat('#,##0');
  if (lastRow > 2) {
    sheet.getRange(`E3:E${lastRow}`).setNumberFormat('+#,##0;-#,##0;0');
    sheet.getRange(`F3:F${lastRow}`).setNumberFormat('0.0%');
  }
  if (lastRow > 13) {
    sheet.getRange(`G14:G${lastRow}`).setNumberFormat('0.0%');
  }

  // チャート作成（売上 + 前年同月比の複合グラフ）
  removeAllCharts_(sheet);

  if (lastRow > 13) {
    // 前年同月比の折れ線グラフ（棒グラフから折れ線に変更）
    // 列構成: A=年月, B=年, C=月, D=売上, E=前月比, F=前月比率, G=前年同月比
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange('A1:A' + lastRow))  // X軸: 年月
      .addRange(sheet.getRange('G1:G' + lastRow))  // Y軸: 前年同月比
      .setPosition(2, 9, 0, 0)
      .setOption('title', '前年同月比推移')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('vAxis', { title: '前年同月比 (%)' })
      .setOption('series', {
        0: { color: theme.primary, lineWidth: 3, pointSize: 6 }
      });

    sheet.insertChart(chartBuilder.build());
  } else if (lastRow > 1) {
    // データが少ない場合は前月比率を表示
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange('A1:A' + lastRow))  // X軸: 年月
      .addRange(sheet.getRange('F1:F' + lastRow))  // Y軸: 前月比率
      .setPosition(2, 7, 0, 0)
      .setOption('title', '前月比率推移')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('vAxis', { title: '前月比率 (%)' })
      .setOption('series', {
        0: { color: theme.primary, lineWidth: 3, pointSize: 6 }
      });

    sheet.insertChart(chartBuilder.build());
  }
}

// ========================================
// スライド生成（既存コードを流用）
// ========================================

function generateOrUpdateSlide(params) {
  try {
    const config = getScriptProperties_();
    const { periodType, targetDate, forceNew, aiComment } = params;

    // periodTypeに応じてテンプレートIDを取得
    let templateId = '';
    if (periodType === 'yearly' && config.slideTemplateIdYearly) {
      templateId = config.slideTemplateIdYearly;
    } else if (periodType === 'weekly' && config.slideTemplateIdWeekly) {
      templateId = config.slideTemplateIdWeekly;
    } else if (periodType === 'monthly' && config.slideTemplateIdMonthly) {
      templateId = config.slideTemplateIdMonthly;
    } else {
      // フォールバック: デフォルトのテンプレートID
      templateId = config.slideTemplateId;
    }

    if (!templateId) {
      throw new Error('スライドテンプレートIDが設定されていません。初期セットアップを実行してください。');
    }

    const data = getReportData_(periodType, targetDate, aiComment || null);
    const configWithTemplate = { ...config, slideTemplateId: templateId };

    let presentation;
    let isNew = false;

    if (forceNew || !config.currentSlideId) {
      presentation = createNewSlide_(configWithTemplate, data);
      isNew = true;
      saveScriptProperties_({ currentSlideId: presentation.getId() });
    } else {
      try {
        presentation = SlidesApp.openById(config.currentSlideId);
        updateSlide_(presentation, data, configWithTemplate);
      } catch (error) {
        presentation = createNewSlide_(configWithTemplate, data);
        isNew = true;
        saveScriptProperties_({ currentSlideId: presentation.getId() });
      }
    }

    return {
      success: true,
      message: isNew ? '新しいスライドを作成しました' : 'スライドを更新しました',
      slideId: presentation.getId(),
      slideUrl: presentation.getUrl(),
      isNew: isNew
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

function createNewSlide_(config, data) {
  const template = DriveApp.getFileById(config.slideTemplateId);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `${data.reportTitle}_${data.period}_${timestamp}`;

  let newFile;
  if (config.outputFolderId) {
    const folder = DriveApp.getFolderById(config.outputFolderId);
    newFile = template.makeCopy(fileName, folder);
  } else {
    newFile = template.makeCopy(fileName);
  }

  const presentation = SlidesApp.openById(newFile.getId());
  applyDataToSlide_(presentation, data);
  return presentation;
}

function updateSlide_(presentation, data, config) {
  const slides = presentation.getSlides();
  
  // テンプレートから新しいスライドを追加（表紙以外）
  if (config && config.slideTemplateId) {
    try {
      // テンプレートファイルを一時的にコピー
      const templateFile = DriveApp.getFileById(config.slideTemplateId);
      const tempTemplate = templateFile.makeCopy('temp_template_' + Date.now());
      const tempPresentation = SlidesApp.openById(tempTemplate.getId());
      const templateSlides = tempPresentation.getSlides();
      
      // 既存のスライドを削除（表紙以外）
      for (let i = slides.length - 1; i > 0; i--) {
        slides[i].remove();
      }
      
      // テンプレートからスライドをコピー（表紙以外）
      // Google Slides APIの制限により、既存のスライドを直接コピーできないため、
      // テンプレートファイルを一時的にコピーして、そこからスライドを取得
      for (let i = 1; i < templateSlides.length; i++) {
        const templateSlide = templateSlides[i];
        const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        
        // テンプレートスライドの背景をコピー
        try {
          const templateShapes = templateSlide.getShapes();
          templateShapes.forEach(shape => {
            try {
              const shapeType = shape.getShapeType();
              const left = shape.getLeft();
              const top = shape.getTop();
              const width = shape.getWidth();
              const height = shape.getHeight();
              
              if (shapeType === SlidesApp.ShapeType.TEXT_BOX) {
                const text = shape.getText().asString();
                const newShape = newSlide.insertTextBox(text, left, top, width, height);
                const textStyle = shape.getText().getTextStyle();
                const newTextStyle = newShape.getText().getTextStyle();
                if (textStyle.getFontSize()) newTextStyle.setFontSize(textStyle.getFontSize());
                if (textStyle.isBold()) newTextStyle.setBold(true);
                if (textStyle.isItalic()) newTextStyle.setItalic(true);
                if (textStyle.getForegroundColor()) {
                  newTextStyle.setForegroundColor(textStyle.getForegroundColor());
                }
                if (textStyle.getFontFamily()) {
                  newTextStyle.setFontFamily(textStyle.getFontFamily());
                }
              } else {
                const newShape = newSlide.insertShape(shapeType, left, top, width, height);
                try {
                  const fill = shape.getFill();
                  if (fill && fill.getSolidFill) {
                    const color = fill.getSolidFill().getColor();
                    newShape.getFill().setSolidFill(color);
                  }
                } catch (e) {
                  // フィル設定に失敗しても続行
                }
                try {
                  if (shape.getBorder()) {
                    newShape.getBorder().setTransparent();
                  }
                } catch (e) {
                  // ボーダー設定に失敗しても続行
                }
              }
            } catch (e) {
              Logger.log('シェイプコピーエラー: ' + e);
            }
          });
        } catch (e) {
          Logger.log('スライドコピーエラー: ' + e);
        }
      }
      
      // 一時テンプレートを削除
      tempTemplate.setTrashed(true);
    } catch (e) {
      Logger.log('テンプレートからのスライドコピーエラー: ' + e);
      // エラー時は既存の方法で続行
      for (let i = slides.length - 1; i > 0; i--) {
        slides[i].remove();
      }
    }
  } else {
    // テンプレートがない場合は既存の方法
    for (let i = slides.length - 1; i > 0; i--) {
      slides[i].remove();
    }
  }
  
  applyDataToSlide_(presentation, data);
}

function applyDataToSlide_(presentation, data) {
  // 成長率ラベルを期間タイプに応じて設定
  let growthRateLabel = '成長率';
  if (data.periodType === 'monthly') {
    growthRateLabel = '前月比';
  } else if (data.periodType === 'yearly') {
    growthRateLabel = '前年比';
  } else if (data.periodType === 'weekly') {
    growthRateLabel = '前週比';
  }

  // タイトルと期間を組み合わせる
  const combinedTitle = data.period ? `${data.period} ${data.reportTitle}` : data.reportTitle;

  const replacements = {
    '{{reportTitle}}': combinedTitle,
    '{{period}}': data.period,
    '{{generatedAt}}': Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
    '{{totalSales}}': formatNumber_(data.totalSales),
    '{{totalSalesChange}}': formatPercent_(data.totalSalesChange),
    '{{growthRateLabel}}': growthRateLabel,
    '{{topRegion}}': data.topRegion,
    '{{topRegionSales}}': formatNumber_(data.topRegionSales),
    '{{topPerson}}': data.topPerson,
    '{{topPersonSales}}': formatNumber_(data.topPersonSales),
    '{{aiComment}}': data.aiComment || ''
  };

  const slides = presentation.getSlides();

  // 表紙（1枚目）のテキストボックスを個別に処理
  if (slides.length > 0) {
    try {
      const slide1 = slides[0];
      const shapes = slide1.getShapes();
      let titleUpdated = false;

      // すべてのテキストボックスをチェックしてタイトルを更新
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
          const text = shape.getText().asString().trim();
          
          // {{reportTitle}}プレースホルダーがある場合は優先して更新
          if (text.includes('{{reportTitle}}')) {
            shape.getText().setText(combinedTitle);
            titleUpdated = true;
            Logger.log('タイトル更新（{{reportTitle}}）: ' + combinedTitle);
            continue;
          }
          
          // 「月次」「年次」「週次」を含むテキストボックスをタイトルとして更新
          if ((text.includes('月次') || text.includes('年次') || text.includes('週次')) && text.includes('レポート')) {
            shape.getText().setText(combinedTitle);
            titleUpdated = true;
            Logger.log('タイトル更新（月次/年次/週次検出）: ' + combinedTitle);
            continue;
          }
          
          // 大きなフォントサイズ（40以上）で「レポート」を含むテキストボックスをタイトルとして更新
          try {
            const fontSize = shape.getText().getTextStyle().getFontSize();
            if (fontSize >= 40 && text.includes('レポート')) {
              shape.getText().setText(combinedTitle);
              titleUpdated = true;
              Logger.log('タイトル更新（大きなフォントサイズ）: ' + combinedTitle);
              continue;
            }
          } catch (e) {
            // フォントサイズ取得エラーは無視
          }
          
          // {{period}}プレースホルダーがある場合は空にする
          if (text.includes('{{period}}')) {
            shape.getText().setText('');
            Logger.log('{{period}}を削除');
          }
          
          // {{generatedAt}}プレースホルダーを処理
          if (text.includes('{{generatedAt}}')) {
            shape.getText().setText('Generated at ' + replacements['{{generatedAt}}']);
          }
        }
      }

      // タイトルが更新されなかった場合、最大フォントサイズのテキストボックスをタイトルとして更新
      if (!titleUpdated) {
        let largestTextBox = null;
        let largestFontSize = 0;
        
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
            const text = shape.getText().asString().trim();
            if (text === '') continue;
            
            try {
              const fontSize = shape.getText().getTextStyle().getFontSize();
              if (fontSize > largestFontSize) {
                largestFontSize = fontSize;
                largestTextBox = shape;
              }
            } catch (e) {
              // フォントサイズ取得エラーは無視
            }
          }
        }
        
        if (largestTextBox && largestFontSize >= 30) {
          largestTextBox.getText().setText(combinedTitle);
          titleUpdated = true;
          Logger.log('タイトル更新（最大フォントサイズ）: ' + combinedTitle);
        }
      }

      // 最後の手段：すべてのテキストボックスを再チェックして確実に更新
      if (!titleUpdated) {
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
            const text = shape.getText().asString();
            // 「月次」「年次」「週次」のいずれかを含むテキストボックスをすべて更新
            if (text.includes('月次') || text.includes('年次') || text.includes('週次')) {
              // 期間部分（例：「2000年01月」）を含む可能性があるので、テキストボックス全体を置換
              shape.getText().setText(combinedTitle);
              titleUpdated = true;
              Logger.log('タイトル更新（最終手段）: ' + combinedTitle);
              break;
            }
          }
        }
      }
      
      // さらに確実にするため、replaceAllTextでも置換（期間部分を含むパターンも含む）
      slide1.replaceAllText('{{reportTitle}}', combinedTitle);
      // 既存のタイトルパターンを置換（期間部分は残して、レポートタイプ部分だけを更新）
      if (data.periodType === 'yearly') {
        slide1.replaceAllText('月次売上レポート', '年次売上レポート');
        slide1.replaceAllText('週次売上レポート', '年次売上レポート');
        slide1.replaceAllText('月次売上 レポート', '年次売上レポート');
        slide1.replaceAllText('週次売上 レポート', '年次売上レポート');
      } else if (data.periodType === 'weekly') {
        slide1.replaceAllText('月次売上レポート', '週次売上レポート');
        slide1.replaceAllText('年次売上レポート', '週次売上レポート');
        slide1.replaceAllText('月次売上 レポート', '週次売上レポート');
        slide1.replaceAllText('年次売上 レポート', '週次売上レポート');
      } else {
        slide1.replaceAllText('年次売上レポート', '月次売上レポート');
        slide1.replaceAllText('週次売上レポート', '月次売上レポート');
        slide1.replaceAllText('年次売上 レポート', '月次売上レポート');
        slide1.replaceAllText('週次売上 レポート', '月次売上レポート');
      }
    } catch (e) {
      Logger.log('表紙テキスト設定エラー: ' + e);
    }
  }

  // テキスト置換を実行（表紙以外も含む）
  // {{reportTitle}}は表紙で個別処理済みなので、他のスライドでのみ置換
  Object.keys(replacements).forEach(key => {
    const value = replacements[key];
    if (key === '{{reportTitle}}') {
      // {{reportTitle}}は表紙で個別処理済みなので、他のスライドのみ置換
      // ただし、表紙以外に{{reportTitle}}がある場合は置換
      if (slides.length > 1) {
        for (let i = 1; i < slides.length; i++) {
          slides[i].replaceAllText(key, value);
        }
      }
    } else if (value) {
      presentation.replaceAllText(key, value);
    } else {
      // 空の場合はプレースホルダーを削除
      presentation.replaceAllText(key, '');
    }
  });

  insertChartsFromSheet_(presentation);
}

/**
 * マークダウン風のテキストをGoogle Slides用にフォーマット
 */
function formatMarkdownLikeText_(text) {
  if (!text) return '';
  
  // マークダウンの見出し記号を削除して改行に変換
  let formatted = text
    .replace(/^#{1,6}\s+/gm, '')  // # 見出しを削除
    .replace(/\*\*(.*?)\*\*/g, '$1')  // **太字** を通常テキストに
    .replace(/\*(.*?)\*/g, '$1')     // *斜体* を通常テキストに
    .replace(/^\*\s+/gm, '・ ')      // * リストを・に変換
    .replace(/^-\s+/gm, '・ ')       // - リストを・に変換
    .replace(/^\d+\.\s+/gm, '')     // 番号リストを削除
    .replace(/\[([^\]]+)\]\([^\)]+\)/g, '$1')  // [リンク](URL) をリンクテキストに
    .replace(/\n{3,}/g, '\n\n')      // 3つ以上の連続改行を2つに
    .trim();
  
  return formatted;
}

function insertChartsFromSheet_(presentation) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slides = presentation.getSlides();

  // Slide 3: 地域別売上グラフ（16:9スライドに収まるように調整、位置を上に）
  const regionalSheet = ss.getSheetByName('RegionalSales');
  if (regionalSheet && regionalSheet.getCharts().length > 0 && slides.length > 2) {
    // 既存のグラフを削除
    const existingCharts = slides[2].getSheetsCharts();
    existingCharts.forEach(chart => chart.remove());

    const chart = regionalSheet.getCharts()[0];
    // Y座標を120から80に変更してグラフを上に移動（大きさは600x330のまま）
    slides[2].insertSheetsChart(chart, 60, 80, 600, 330);
  }

  // Slide 4: 担当者別売上グラフ（16:9スライドに収まるように調整、位置を上に）
  const personSheet = ss.getSheetByName('PersonSales');
  if (personSheet && personSheet.getCharts().length > 0 && slides.length > 3) {
    // 既存のグラフを削除
    const existingCharts = slides[3].getSheetsCharts();
    existingCharts.forEach(chart => chart.remove());

    const chart = personSheet.getCharts()[0];
    // Y座標を120から80に変更してグラフを上に移動（大きさは600x330のまま）
    slides[3].insertSheetsChart(chart, 60, 80, 600, 330);
  }
}

/**
 * RawSalesDataから期間に基づいてフィルタリングしたデータを取得
 */
function getFilteredRawData_(periodType, targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RawSalesData');

  if (!rawSheet || rawSheet.getLastRow() <= 1) {
    return [];
  }

  const date = targetDate ? new Date(targetDate) : new Date();
  let startDate, endDate;

  // periodTypeに基づいて日付範囲を計算
  switch (periodType) {
    case 'monthly':
      startDate = new Date(date.getFullYear(), date.getMonth(), 1);
      endDate = new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59);
      break;
    case 'weekly':
      // 週の始まり（月曜日）を計算
      const dayOfWeek = date.getDay();
      const daysToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      startDate = new Date(date);
      startDate.setDate(startDate.getDate() + daysToMonday);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(startDate);
      endDate.setDate(endDate.getDate() + 6);
      endDate.setHours(23, 59, 59, 999);
      break;
    case 'yearly':
      startDate = new Date(date.getFullYear(), 0, 1);
      endDate = new Date(date.getFullYear(), 11, 31, 23, 59, 59);
      break;
    default:
      startDate = new Date(date);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(date);
      endDate.setHours(23, 59, 59, 999);
  }

  // データを取得
  const data = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, rawSheet.getLastColumn()).getValues();
  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];

  // 日付でフィルタリング
  const filteredData = data
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    })
    .filter(row => {
      const rowDate = row.Date || row['Date'];
      if (!rowDate) return false;
      const dateValue = rowDate instanceof Date ? rowDate : new Date(rowDate);
      return dateValue >= startDate && dateValue <= endDate;
    });

  return filteredData;
}

/**
 * フィルタリング後のデータから集計シートを作成（一時的）
 */
function createTemporaryAggregationSheets_(filteredRawData) {
  // データが空の場合は空の配列を返す
  if (!filteredRawData || filteredRawData.length === 0) {
    return { regionalData: [], personData: [] };
  }
  
  // 集計データを作成
  const regionalData = [];
  const personData = [];
  
  // 地域別集計
  const regionMap = {};
  filteredRawData.forEach(row => {
    const region = row.Region || row['Region'] || 'N/A';
    const sales = parseFloat(row.TotalSales || row['TotalSales'] || 0);
    if (!regionMap[region]) {
      regionMap[region] = 0;
    }
    regionMap[region] += sales;
  });
  
  Object.keys(regionMap).forEach(region => {
    regionalData.push({
      '地域': region,
      '売上': regionMap[region],
      Region: region,
      Sales: regionMap[region]
    });
  });
  
  // 担当者別集計
  const personMap = {};
  filteredRawData.forEach(row => {
    const person = row.Person || row['Person'] || 'N/A';
    const sales = parseFloat(row.TotalSales || row['TotalSales'] || 0);
    if (!personMap[person]) {
      personMap[person] = { sales: 0, count: 0 };
    }
    personMap[person].sales += sales;
    personMap[person].count += 1;
  });
  
  Object.keys(personMap).forEach(person => {
    personData.push({
      '担当者': person,
      '売上': personMap[person].sales,
      '件数': personMap[person].count,
      Person: person,
      Sales: personMap[person].sales
    });
  });
  
  return { regionalData, personData };
}

function getReportData_(periodType, targetDate, customAiComment = null) {
  const config = getScriptProperties_();
  const period = formatPeriod_(periodType, targetDate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // FilteredSummaryシートが存在する場合、そこからデータを取得（フィルタリング済み）
  const filteredSummarySheet = ss.getSheetByName('FilteredSummary');
  let totalSales, totalSalesChange, topRegion, topRegionSales, topPerson, topPersonSales;

  if (filteredSummarySheet && filteredSummarySheet.getLastRow() > 2) {
    // FilteredSummaryシートから直接読み取る（1行目：対象期間、2行目：ヘッダー、3行目以降：データ）
    const summaryData = filteredSummarySheet.getRange(3, 1, filteredSummarySheet.getLastRow() - 2, 2).getValues();
    const summaryMap = {};
    summaryData.forEach(row => {
      summaryMap[row[0]] = row[1];
    });

    totalSales = summaryMap['合計売上'] || 0;
    totalSalesChange = summaryMap['前月比'] || summaryMap['前年比'] || summaryMap['前週比'] || summaryMap['成長率'] || 0;
    topRegion = summaryMap['トップ地域'] || 'N/A';
    topRegionSales = summaryMap['トップ地域売上'] || 0;
    topPerson = summaryMap['トップ担当者'] || 'N/A';
    topPersonSales = summaryMap['トップ担当者売上'] || 0;
  } else {
    // FilteredSummaryシートがない場合、現在のシートから取得（フィルタリングなし）
    const regionalData = getSheetData_('RegionalSales');
    const personData = getSheetData_('PersonSales');

    // 売上データを取得
    totalSales = regionalData.reduce((sum, row) => {
      const sales = row['売上'] || row.Sales || 0;
      return sum + (typeof sales === 'number' ? sales : 0);
    }, 0);

    topRegion = regionalData.length > 0 ? (regionalData[0]['地域'] || regionalData[0].Region || 'N/A') : 'N/A';
    topRegionSales = regionalData.length > 0 ? (regionalData[0]['売上'] || regionalData[0].Sales || 0) : 0;
    topPerson = personData.length > 0 ? (personData[0]['担当者'] || personData[0].Person || 'N/A') : 'N/A';
    topPersonSales = personData.length > 0 ? (personData[0]['売上'] || personData[0].Sales || 0) : 0;

    // 成長率を計算（MonthlySalesシートから取得）
    totalSalesChange = 0;
    try {
      const monthlySheet = ss.getSheetByName('MonthlySales');
      if (monthlySheet && monthlySheet.getLastRow() > 1) {
        const date = new Date(targetDate);
        const currentYear = date.getFullYear();
        const currentMonth = date.getMonth() + 1;

        const lastRow = monthlySheet.getLastRow();
        const monthlyData = monthlySheet.getRange(2, 1, lastRow - 1, 7).getValues();

        if (periodType === 'monthly') {
          const targetRow = monthlyData.find(row => row[1] === currentYear && row[2] === currentMonth);
          if (targetRow && targetRow[5] !== '' && targetRow[5] !== '-') {
            totalSalesChange = typeof targetRow[5] === 'number' ? targetRow[5] : 0;
          }
        } else if (periodType === 'yearly') {
          const currentYearRows = monthlyData.filter(row => row[1] === currentYear);
          if (currentYearRows.length > 0) {
            const validRates = currentYearRows
              .map(row => row[6])
              .filter(rate => rate !== '' && rate !== '-' && typeof rate === 'number');
            if (validRates.length > 0) {
              totalSalesChange = validRates.reduce((sum, rate) => sum + rate, 0) / validRates.length;
            }
          }
        }
      }
    } catch (e) {
      Logger.log('成長率計算エラー: ' + e);
      totalSalesChange = 0;
    }
  }

  // AIコメント（UIから渡された値を使用、なければ自動生成）
  let aiComment = customAiComment || '';

  // UIから値が渡されていない場合のみ自動生成
  if (!customAiComment && config.geminiApiKey) {
    try {
      // AIコメントを生成
      const commentResult = generateAICommentForData_({
        totalSales: totalSales,
        totalSalesChange: totalSalesChange,
        topRegion: topRegion,
        topRegionSales: topRegionSales,
        topPerson: topPerson,
        topPersonSales: topPersonSales
      });

      if (commentResult.success) {
        aiComment = commentResult.text;
      }
    } catch (e) {
      Logger.log('AIコメント生成エラー: ' + e);
    }
  }

  // periodTypeに応じてタイトルを動的に生成
  let reportTitle;
  if (periodType === 'yearly') {
    reportTitle = '年次売上レポート';
  } else if (periodType === 'weekly') {
    reportTitle = '週次売上レポート';
  } else {
    reportTitle = '月次売上レポート';
  }

  return {
    reportTitle: reportTitle,
    period: period,
    periodType: periodType,
    totalSales: totalSales,
    totalSalesChange: totalSalesChange,
    topRegion: topRegion,
    topRegionSales: topRegionSales,
    topPerson: topPerson,
    topPersonSales: topPersonSales,
    aiComment: aiComment
  };
}

/**
 * レポートデータをプレビュー（UI用）
 * 期間選択時に自動的に呼び出され、フィルタリング後のデータとAIコメントを返す
 */
/**
 * フィルタリングを実行してプレビューデータを返す
 */
function previewReportData(periodType, targetDate) {
  try {
    // 1. QUERY関数を更新してフィルタリング
    updateQueryFormulasWithFilter_(periodType, targetDate);

    // 2. FilteredSummaryシートを作成して集計データを取得
    const summaryData = createFilteredSummarySheet_(periodType, targetDate);

    // 3. AIコメントを生成
    const config = getScriptProperties_();
    let aiComment = '';

    if (config.geminiApiKey) {
      try {
        const commentResult = generateAICommentForData_({
          totalSales: summaryData.totalSales,
          totalSalesChange: summaryData.totalSalesChange || 0,
          topRegion: summaryData.topRegion,
          topRegionSales: summaryData.topRegionSales,
          topPerson: summaryData.topPerson,
          topPersonSales: summaryData.topPersonSales
        });
        aiComment = commentResult.success ? commentResult.text : '';
      } catch (e) {
        Logger.log('AI生成エラー: ' + e);
      }
    }

    return {
      success: true,
      totalSales: summaryData.totalSales,
      totalSalesChange: summaryData.totalSalesChange,
      topRegion: summaryData.topRegion,
      topRegionSales: summaryData.topRegionSales,
      topPerson: summaryData.topPerson,
      topPersonSales: summaryData.topPersonSales,
      aiComment: aiComment
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * フィルターを解除して全期間のデータに戻す
 */
function clearDataFilter() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // RegionalSalesのQUERY式を元に戻す（WHERE句なし）
    const regionalSheet = ss.getSheetByName('RegionalSales');
    if (regionalSheet) {
      const regionalFormula = '=QUERY(RawSalesData!A:H, "SELECT B, SUM(H) WHERE B IS NOT NULL GROUP BY B ORDER BY SUM(H) DESC LABEL B \'地域\', SUM(H) \'売上\'", 1)';
      regionalSheet.getRange('A1').setFormula(regionalFormula);

      // C列の万円単位計算を更新
      Utilities.sleep(1000);
      const lastRow = regionalSheet.getLastRow();
      if (lastRow > 1) {
        regionalSheet.getRange('C1').setValue('売上(万円)');
        for (let i = 2; i <= lastRow; i++) {
          regionalSheet.getRange(`C${i}`).setFormula(`=B${i}/10000`);
        }
      }
    }

    // PersonSalesのQUERY式を元に戻す（WHERE句なし）
    const personSheet = ss.getSheetByName('PersonSales');
    if (personSheet) {
      const personFormula = '=QUERY(RawSalesData!A:H, "SELECT C, SUM(H), COUNT(H) WHERE C IS NOT NULL GROUP BY C ORDER BY SUM(H) DESC LABEL C \'担当者\', SUM(H) \'売上\', COUNT(H) \'件数\'", 1)';
      personSheet.getRange('A1').setFormula(personFormula);

      // E列の万円単位計算とD列の平均単価を更新
      Utilities.sleep(1000);
      const lastRow = personSheet.getLastRow();
      if (lastRow > 1) {
        personSheet.getRange('D1').setValue('平均単価');
        personSheet.getRange('E1').setValue('売上(万円)');
        for (let i = 2; i <= lastRow; i++) {
          personSheet.getRange(`D${i}`).setFormula(`=IF(C${i}>0, B${i}/C${i}, 0)`);
          personSheet.getRange(`E${i}`).setFormula(`=B${i}/10000`);
        }
      }
    }

    // FilteredSummaryシートを削除
    const summarySheet = ss.getSheetByName('FilteredSummary');
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }

    return {
      success: true,
      message: 'フィルターを解除しました'
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * QUERY関数にWHERE句を追加してフィルタリング
 */
function updateQueryFormulasWithFilter_(periodType, targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const date = new Date(targetDate);

  // 日付範囲の条件文字列を作成
  let whereClause = '';

  if (periodType === 'monthly') {
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    whereClause = `WHERE B IS NOT NULL AND YEAR(A) = ${year} AND MONTH(A) = ${month}`;
  } else if (periodType === 'weekly') {
    const startDate = new Date(date);
    const dayOfWeek = date.getDay();
    const daysToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    startDate.setDate(startDate.getDate() + daysToMonday);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 6);

    const startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    whereClause = `WHERE B IS NOT NULL AND A >= date '${startStr}' AND A <= date '${endStr}'`;
  } else if (periodType === 'yearly') {
    const year = date.getFullYear();
    whereClause = `WHERE B IS NOT NULL AND YEAR(A) = ${year}`;
  } else {
    whereClause = 'WHERE B IS NOT NULL';
  }

  // RegionalSalesのQUERY式を更新
  const regionalSheet = ss.getSheetByName('RegionalSales');
  if (regionalSheet) {
    const regionalFormula = `=QUERY(RawSalesData!A:H, "SELECT B, SUM(H) ${whereClause} GROUP BY B ORDER BY SUM(H) DESC LABEL B '地域', SUM(H) '売上'", 1)`;
    regionalSheet.getRange('A1').setFormula(regionalFormula);

    // C列の万円単位計算を更新
    Utilities.sleep(1000);
    const lastRow = regionalSheet.getLastRow();
    if (lastRow > 1) {
      regionalSheet.getRange('C1').setValue('売上(万円)');
      for (let i = 2; i <= lastRow; i++) {
        regionalSheet.getRange(`C${i}`).setFormula(`=B${i}/10000`);
      }
    }
  }

  // PersonSalesのQUERY式を更新（C列：担当者）
  const personSheet = ss.getSheetByName('PersonSales');
  if (personSheet) {
    const personFormula = `=QUERY(RawSalesData!A:H, "SELECT C, SUM(H), COUNT(H) ${whereClause} GROUP BY C ORDER BY SUM(H) DESC LABEL C '担当者', SUM(H) '売上', COUNT(H) '件数'", 1)`;
    personSheet.getRange('A1').setFormula(personFormula);

    // E列の万円単位計算とD列の平均単価を更新
    Utilities.sleep(1000);
    const lastRow = personSheet.getLastRow();
    if (lastRow > 1) {
      personSheet.getRange('D1').setValue('平均単価');
      personSheet.getRange('E1').setValue('売上(万円)');
      for (let i = 2; i <= lastRow; i++) {
        personSheet.getRange(`D${i}`).setFormula(`=IF(C${i}>0, B${i}/C${i}, 0)`);
        personSheet.getRange(`E${i}`).setFormula(`=B${i}/10000`);
      }
    }
  }
}

/**
 * フィルタリング結果のサマリーシートを作成
 */
function createFilteredSummarySheet_(periodType, targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const date = new Date(targetDate);

  // FilteredSummaryシートを作成（既存は削除）
  let summarySheet = ss.getSheetByName('FilteredSummary');
  if (summarySheet) ss.deleteSheet(summarySheet);
  summarySheet = ss.insertSheet('FilteredSummary');

  // 計算待機
  Utilities.sleep(1500);

  // RegionalSalesから集計
  const regionalSheet = ss.getSheetByName('RegionalSales');
  let totalSales = 0;
  let topRegion = 'N/A';
  let topRegionSales = 0;

  if (regionalSheet && regionalSheet.getLastRow() > 1) {
    const regionalData = regionalSheet.getRange(2, 1, regionalSheet.getLastRow() - 1, 2).getValues();
    totalSales = regionalData.reduce((sum, row) => sum + (row[1] || 0), 0);

    if (regionalData.length > 0) {
      topRegion = regionalData[0][0] || 'N/A';
      topRegionSales = regionalData[0][1] || 0;
    }
  }

  // PersonSalesから集計
  const personSheet = ss.getSheetByName('PersonSales');
  let topPerson = 'N/A';
  let topPersonSales = 0;

  if (personSheet && personSheet.getLastRow() > 1) {
    const personData = personSheet.getRange(2, 1, 1, 2).getValues();
    if (personData.length > 0) {
      topPerson = personData[0][0] || 'N/A';
      topPersonSales = personData[0][1] || 0;
    }
  }

  // 成長率を計算（期間タイプに応じて前月比または前年比）
  // MonthlySalesシートから直接F列（前月比率）またはG列（前年同月比）を読み取る
  let totalSalesChange = 0;
  try {
    const monthlySheet = ss.getSheetByName('MonthlySales');
    if (monthlySheet && monthlySheet.getLastRow() > 1) {
      const currentYear = date.getFullYear();
      const currentMonth = date.getMonth() + 1;

      // A列（年月）、B列（年）、C列（月）、D列（売上）、F列（前月比率）、G列（前年同月比）を取得
      const lastRow = monthlySheet.getLastRow();
      const monthlyData = monthlySheet.getRange(2, 1, lastRow - 1, 7).getValues();

      if (periodType === 'monthly') {
        // 月次：F列（前月比率）を取得
        const targetRow = monthlyData.find(row => {
          const year = row[1];  // B列（年）
          const month = row[2]; // C列（月）
          return year === currentYear && month === currentMonth;
        });

        if (targetRow && targetRow[5] !== '' && targetRow[5] !== '-') {
          // F列（前月比率）はインデックス5
          totalSalesChange = typeof targetRow[5] === 'number' ? targetRow[5] : 0;
        }
      } else if (periodType === 'yearly') {
        // 年次：G列（前年同月比）の平均を計算
        const currentYearRows = monthlyData.filter(row => row[1] === currentYear);

        if (currentYearRows.length > 0) {
          // 前年同月比がある行のみ抽出して平均
          const validRates = currentYearRows
            .map(row => row[6]) // G列（前年同月比）はインデックス6
            .filter(rate => rate !== '' && rate !== '-' && typeof rate === 'number');

          if (validRates.length > 0) {
            totalSalesChange = validRates.reduce((sum, rate) => sum + rate, 0) / validRates.length;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('成長率計算エラー: ' + e);
    totalSalesChange = 0;
  }

  // 成長率のラベルを期間タイプに応じて設定
  let growthRateLabel = '成長率';
  if (periodType === 'monthly') {
    growthRateLabel = '前月比';
  } else if (periodType === 'yearly') {
    growthRateLabel = '前年比';
  } else if (periodType === 'weekly') {
    growthRateLabel = '前週比';
  }

  // 対象期間を文字列化
  let periodLabel = '';
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  if (periodType === 'monthly') {
    periodLabel = `${year}年${month}月`;
  } else if (periodType === 'yearly') {
    periodLabel = `${year}年`;
  } else if (periodType === 'weekly') {
    const weekNum = Math.ceil(date.getDate() / 7);
    periodLabel = `${year}年${month}月 第${weekNum}週`;
  }

  // サマリーシートに書き込み（対象期間を追加）
  summarySheet.getRange('A1').setValue('対象期間');
  summarySheet.getRange('B1').setValue(periodLabel);
  summarySheet.getRange('A2').setValue('項目');
  summarySheet.getRange('B2').setValue('値');
  summarySheet.getRange('A3').setValue('合計売上');
  summarySheet.getRange('B3').setValue(totalSales);
  summarySheet.getRange('A4').setValue(growthRateLabel);
  summarySheet.getRange('B4').setValue(totalSalesChange);
  summarySheet.getRange('A5').setValue('トップ地域');
  summarySheet.getRange('B5').setValue(topRegion);
  summarySheet.getRange('A6').setValue('トップ地域売上');
  summarySheet.getRange('B6').setValue(topRegionSales);
  summarySheet.getRange('A7').setValue('トップ担当者');
  summarySheet.getRange('B7').setValue(topPerson);
  summarySheet.getRange('A8').setValue('トップ担当者売上');
  summarySheet.getRange('B8').setValue(topPersonSales);

  // スタイル設定
  summarySheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  summarySheet.getRange('A2:B2').setFontWeight('bold').setBackground('#0f9d58').setFontColor('#ffffff');
  summarySheet.getRange('B3:B3').setNumberFormat('#,##0');
  summarySheet.getRange('B4:B4').setNumberFormat('0.0%');
  summarySheet.getRange('B6:B6').setNumberFormat('#,##0');
  summarySheet.getRange('B8:B8').setNumberFormat('#,##0');
  summarySheet.setColumnWidth(1, 150);
  summarySheet.setColumnWidth(2, 150);

  return {
    totalSales: totalSales,
    topRegion: topRegion,
    topRegionSales: topRegionSales,
    topPerson: topPerson,
    topPersonSales: topPersonSales,
    totalSalesChange: totalSalesChange
  };
}

function getSheetData_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function formatPeriod_(periodType, targetDate) {
  const date = targetDate ? new Date(targetDate) : new Date();

  switch (periodType) {
    case 'monthly':
      return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年MM月');
    case 'weekly':
      const weekNum = Math.ceil((date.getDate() + 6 - date.getDay()) / 7);
      return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年MM月') + ` 第${weekNum}週`;
    case 'yearly':
      return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy年');
    default:
      return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
}

// ========================================
// Gemini AI機能
// ========================================

function generateTextWithGemini(prompt, customPrompt = '') {
  try {
    const config = getScriptProperties_();

    if (!config.geminiApiKey) {
      return {
        success: false,
        message: 'Gemini APIキーが設定されていません。'
      };
    }

    // Gemini 2.0 Flash Lite を使用
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key=${config.geminiApiKey}`;

    // カスタムプロンプトがある場合はシステムプロンプトと組み合わせる
    const systemPrompt = customPrompt || `あなたは営業レポート分析の専門家です。以下のガイドラインに従ってください：
- 簡潔で具体的な分析を提供する
- 数値データに基づいた客観的な評価を行う
- ビジネスインサイトと実行可能な提案を含める
- ポジティブかつ建設的なトーンで記述する`;

    const payload = {
      contents: [{
        parts: [{
          text: systemPrompt + '\n\n' + prompt
        }]
      }],
      generationConfig: {
        temperature: 0.7,
        topK: 40,
        topP: 0.95,
        maxOutputTokens: 1024
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates.length > 0) {
      const text = json.candidates[0].content.parts[0].text;
      return {
        success: true,
        text: text
      };
    } else {
      return {
        success: false,
        message: 'テキスト生成に失敗しました: ' + (json.error ? json.error.message : '不明なエラー')
      };
    }
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * データからAIコメントを生成（内部関数）
 */
function generateAICommentForData_(data) {
  const prompt = `以下の売上データを分析して、経営陣向けの簡潔なサマリーコメント（150文字以内）を日本語で生成してください:

合計売上: ${formatNumber_(data.totalSales)}
成長率: ${formatPercent_(data.totalSalesChange)}
トップ地域: ${data.topRegion} (${formatNumber_(data.topRegionSales)})
トップ担当者: ${data.topPerson} (${formatNumber_(data.topPersonSales)})

重要な数値を含め、ポジティブで前向きなコメントをお願いします。`;

  return generateTextWithGemini(prompt);
}

/**
 * UI用のAIコメント生成関数（期間指定対応）
 */
function generateAIComment(params) {
  try {
    const { periodType, targetDate } = params || { periodType: 'monthly', targetDate: null };
    const data = getReportData_(periodType, targetDate);
    return generateAICommentForData_({
      totalSales: data.totalSales,
      totalSalesChange: data.totalSalesChange,
      topRegion: data.topRegion,
      topRegionSales: data.topRegionSales,
      topPerson: data.topPerson,
      topPersonSales: data.topPersonSales
    });
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * データからAIインサイトを生成（内部関数）
 */
function generateAIInsightForData_(data) {
  // データを整形
  const regionSummary = data.regionalData.slice(0, 5).map(r => {
    const region = r['地域'] || r.Region || 'N/A';
    const sales = r['売上'] || r.Sales || 0;
    return `${region}: ${formatNumber_(sales)}`;
  }).join(', ');
  
  const personSummary = data.personData.slice(0, 5).map(p => {
    const person = p['担当者'] || p.Person || 'N/A';
    const sales = p['売上'] || p.Sales || 0;
    return `${person}: ${formatNumber_(sales)}`;
  }).join(', ');

  const prompt = `以下の売上データから、ビジネスインサイトと具体的なアクションプラン（200文字程度）を日本語で生成してください:

【概要】
合計売上: ${formatNumber_(data.totalSales)}
成長率: ${formatPercent_(data.totalSalesChange)}

【地域別トップ5】
${regionSummary || 'データなし'}

【担当者別トップ5】
${personSummary || 'データなし'}

データから読み取れる課題や機会、次に取るべき具体的なアクションを提案してください。`;

  return generateTextWithGemini(prompt);
}

/**
 * UI用のAIインサイト生成関数
 */
function generateAIInsight() {
  try {
    const data = getReportData_('monthly', null);
    const regionalData = getSheetData_('RegionalSales');
    const personData = getSheetData_('PersonSales');

    return generateAIInsightForData_({
      totalSales: data.totalSales,
      totalSalesChange: data.totalSalesChange,
      regionalData: regionalData,
      personData: personData
    });
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

function generateCustomText(userPrompt, systemPrompt = '') {
  return generateTextWithGemini(userPrompt, systemPrompt);
}

/**
 * Gemini接続テスト
 */
function testGeminiConnection() {
  try {
    const config = getScriptProperties_();

    if (!config.geminiApiKey) {
      return {
        success: false,
        message: 'Gemini APIキーが設定されていません。'
      };
    }

    const testPrompt = 'こんにちは！接続テストです。「接続成功」と日本語で返答してください。';
    const result = generateTextWithGemini(testPrompt, '');

    if (result.success) {
      return {
        success: true,
        message: 'Gemini 2.0 Flash Lite との接続に成功しました！',
        response: result.text
      };
    } else {
      return result;
    }
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

// ========================================
// ユーティリティ
// ========================================

function formatNumber_(num) {
  if (typeof num !== 'number') {
    num = Number(num);
  }
  if (isNaN(num)) {
    return '¥0';
  }
  return '¥' + Math.round(num).toLocaleString('ja-JP');
}

function formatPercent_(num) {
  if (typeof num !== 'number') {
    num = Number(num);
  }
  if (isNaN(num)) {
    return '0%';
  }
  const sign = num >= 0 ? '+' : '';
  return sign + (num * 100).toFixed(1) + '%';
}

function resetCurrentSlideId() {
  try {
    saveScriptProperties_({ currentSlideId: '' });
    return { success: true, message: 'スライドIDをリセットしました' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
/**
 * RawSalesDataから日付範囲と利用可能な期間を取得
 */
function getAvailablePeriods() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('RawSalesData');

    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      return {
        success: false,
        message: 'RawSalesDataが見つかりません。データをインポートしてください。'
      };
    }

    // A列（Date列）のデータを取得
    const dateRange = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 1);
    const dates = dateRange.getValues()
      .map(row => row[0])
      .filter(date => date && date !== '');

    if (dates.length === 0) {
      return {
        success: false,
        message: '日付データが見つかりません'
      };
    }

    // 日付を文字列からDateオブジェクトに変換
    const parsedDates = dates.map(d => {
      if (d instanceof Date) return d;
      return new Date(d);
    }).filter(d => !isNaN(d.getTime()));

    // 最小・最大日付を取得
    const minDate = new Date(Math.min(...parsedDates));
    const maxDate = new Date(Math.max(...parsedDates));

    // 利用可能な年のリストを生成
    const minYear = minDate.getFullYear();
    const maxYear = maxDate.getFullYear();
    const years = [];
    for (let y = minYear; y <= maxYear; y++) {
      years.push(y);
    }

    // 利用可能な年月のリストを生成
    const yearMonths = [];
    for (let y = minYear; y <= maxYear; y++) {
      const startMonth = (y === minYear) ? minDate.getMonth() + 1 : 1;
      const endMonth = (y === maxYear) ? maxDate.getMonth() + 1 : 12;
      for (let m = startMonth; m <= endMonth; m++) {
        yearMonths.push({ year: y, month: m });
      }
    }

    // 利用可能な週のリストを生成（月曜始まり）
    const weeks = [];
    let currentDate = new Date(minDate);
    // 週の始まり（月曜日）に調整
    const dayOfWeek = currentDate.getDay();
    const daysToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    currentDate.setDate(currentDate.getDate() + daysToMonday);

    while (currentDate <= maxDate) {
      const weekEnd = new Date(currentDate);
      weekEnd.setDate(weekEnd.getDate() + 6);

      weeks.push({
        year: currentDate.getFullYear(),
        startDate: Utilities.formatDate(currentDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
        endDate: Utilities.formatDate(weekEnd, 'Asia/Tokyo', 'yyyy-MM-dd'),
        label: formatDateRange_(currentDate, weekEnd)
      });

      currentDate.setDate(currentDate.getDate() + 7);
    }

    return {
      success: true,
      minDate: Utilities.formatDate(minDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      maxDate: Utilities.formatDate(maxDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      years: years,
      yearMonths: yearMonths,
      weeks: weeks
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * 日付範囲をフォーマット（例: 12/1-12/7）
 */
function formatDateRange_(startDate, endDate) {
  const startMonth = startDate.getMonth() + 1;
  const startDay = startDate.getDate();
  const endMonth = endDate.getMonth() + 1;
  const endDay = endDate.getDate();

  if (startMonth === endMonth) {
    return `${startMonth}/${startDay}-${endDay}`;
  } else {
    return `${startMonth}/${startDay}-${endMonth}/${endDay}`;
  }
}
