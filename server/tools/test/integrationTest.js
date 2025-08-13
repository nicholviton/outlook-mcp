// Integration test to verify all modules can be imported and basic functionality works

async function testModuleImports() {
  console.log('🔍 Testing module imports...\n');
  
  const results = [];
  
  // Test common utilities
  console.log('📦 Testing common utilities...');
  try {
    const { clearStylingCache, clearSignatureCache, getStylingCacheStats, applyUserStyling } = await import('../common/sharedUtils.js');
    console.log('  ✅ sharedUtils.js imported successfully');
    
    // Test cache functions
    clearStylingCache();
    clearSignatureCache();
    const stats = getStylingCacheStats();
    console.log('  ✅ Cache functions work correctly');
    
    results.push({ module: 'common/sharedUtils.js', status: 'success' });
  } catch (error) {
    console.log('  ❌ sharedUtils.js import failed:', error.message);
    results.push({ module: 'common/sharedUtils.js', status: 'failed', error: error.message });
  }
  
  try {
    const { getRateLimitMetricsTool, resetRateLimitMetricsTool } = await import('../common/rateLimitUtils.js');
    console.log('  ✅ rateLimitUtils.js imported successfully');
    results.push({ module: 'common/rateLimitUtils.js', status: 'success' });
  } catch (error) {
    console.log('  ❌ rateLimitUtils.js import failed:', error.message);
    results.push({ module: 'common/rateLimitUtils.js', status: 'failed', error: error.message });
  }
  
  // Test email modules
  console.log('\n📧 Testing email modules...');
  const emailModules = [
    { path: '../email/listEmails.js', functions: ['listEmailsTool', 'getEmailTool'] },
    { path: '../email/sendEmail.js', functions: ['sendEmailTool'] },
    { path: '../email/searchEmails.js', functions: ['searchEmailsTool'] },
    { path: '../email/createDraft.js', functions: ['createDraftTool'] },
    { path: '../email/replyEmail.js', functions: ['replyToEmailTool', 'replyAllTool'] },
    { path: '../email/forwardEmail.js', functions: ['forwardEmailTool'] },
    { path: '../email/emailManagement.js', functions: ['deleteEmailTool', 'moveEmailTool', 'markAsReadTool', 'flagEmailTool', 'categorizeEmailTool', 'archiveEmailTool', 'batchProcessEmailsTool'] }
  ];
  
  for (const module of emailModules) {
    try {
      const imports = await import(module.path);
      const missingFunctions = module.functions.filter(fn => !imports[fn]);
      if (missingFunctions.length > 0) {
        throw new Error(`Missing functions: ${missingFunctions.join(', ')}`);
      }
      console.log(`  ✅ ${module.path} imported successfully (${module.functions.length} functions)`);
      results.push({ module: module.path, status: 'success', functions: module.functions.length });
    } catch (error) {
      console.log(`  ❌ ${module.path} import failed:`, error.message);
      results.push({ module: module.path, status: 'failed', error: error.message });
    }
  }
  
  // Test calendar modules
  console.log('\n📅 Testing calendar modules...');
  const calendarModules = [
    { path: '../calendar/listEvents.js', functions: ['listEventsTool'] },
    { path: '../calendar/createEvent.js', functions: ['createEventTool'] },
    { path: '../calendar/eventManagement.js', functions: ['getEventTool', 'updateEventTool', 'deleteEventTool', 'respondToInviteTool', 'validateEventDateTimesTool'] },
    { path: '../calendar/calendarUtils.js', functions: ['createRecurringEventTool', 'findMeetingTimesTool', 'checkAvailabilityTool', 'scheduleOnlineMeetingTool', 'listCalendarsTool', 'getCalendarViewTool', 'getBusyTimesTool', 'buildRecurrencePatternTool', 'createRecurrenceHelperTool', 'checkCalendarPermissionsTool'] }
  ];
  
  for (const module of calendarModules) {
    try {
      const imports = await import(module.path);
      const missingFunctions = module.functions.filter(fn => !imports[fn]);
      if (missingFunctions.length > 0) {
        throw new Error(`Missing functions: ${missingFunctions.join(', ')}`);
      }
      console.log(`  ✅ ${module.path} imported successfully (${module.functions.length} functions)`);
      results.push({ module: module.path, status: 'success', functions: module.functions.length });
    } catch (error) {
      console.log(`  ❌ ${module.path} import failed:`, error.message);
      results.push({ module: module.path, status: 'failed', error: error.message });
    }
  }
  
  // Test folder modules
  console.log('\n📁 Testing folder modules...');
  const folderModules = [
    { path: '../folders/listFolders.js', functions: ['listFoldersTool'] },
    { path: '../folders/createFolder.js', functions: ['createFolderTool'] },
    { path: '../folders/renameFolder.js', functions: ['renameFolderTool'] },
    { path: '../folders/getFolderStats.js', functions: ['getFolderStatsTool'] }
  ];
  
  for (const module of folderModules) {
    try {
      const imports = await import(module.path);
      const missingFunctions = module.functions.filter(fn => !imports[fn]);
      if (missingFunctions.length > 0) {
        throw new Error(`Missing functions: ${missingFunctions.join(', ')}`);
      }
      console.log(`  ✅ ${module.path} imported successfully (${module.functions.length} functions)`);
      results.push({ module: module.path, status: 'success', functions: module.functions.length });
    } catch (error) {
      console.log(`  ❌ ${module.path} import failed:`, error.message);
      results.push({ module: module.path, status: 'failed', error: error.message });
    }
  }
  
  // Test attachment modules
  console.log('\n📎 Testing attachment modules...');
  const attachmentModules = [
    { path: '../attachments/listAttachments.js', functions: ['listAttachmentsTool'] },
    { path: '../attachments/downloadAttachment.js', functions: ['downloadAttachmentTool'] },
    { path: '../attachments/addAttachment.js', functions: ['addAttachmentTool'] },
    { path: '../attachments/scanAttachments.js', functions: ['scanAttachmentsTool'] }
  ];
  
  for (const module of attachmentModules) {
    try {
      const imports = await import(module.path);
      const missingFunctions = module.functions.filter(fn => !imports[fn]);
      if (missingFunctions.length > 0) {
        throw new Error(`Missing functions: ${missingFunctions.join(', ')}`);
      }
      console.log(`  ✅ ${module.path} imported successfully (${module.functions.length} functions)`);
      results.push({ module: module.path, status: 'success', functions: module.functions.length });
    } catch (error) {
      console.log(`  ❌ ${module.path} import failed:`, error.message);
      results.push({ module: module.path, status: 'failed', error: error.message });
    }
  }
  
  // Test main index barrel export
  console.log('\n📦 Testing main index barrel export...');
  try {
    const allTools = await import('../index.js');
    const toolCount = Object.keys(allTools).length;
    console.log(`  ✅ Main index.js imported successfully (${toolCount} exports)`);
    results.push({ module: 'index.js', status: 'success', functions: toolCount });
  } catch (error) {
    console.log('  ❌ Main index.js import failed:', error.message);
    results.push({ module: 'index.js', status: 'failed', error: error.message });
  }
  
  // Summary
  console.log('\n📊 Integration Test Summary:');
  const successful = results.filter(r => r.status === 'success');
  const failed = results.filter(r => r.status === 'failed');
  const totalFunctions = successful.reduce((sum, r) => sum + (r.functions || 0), 0);
  
  console.log(`  ✅ Successful imports: ${successful.length}/${results.length}`);
  console.log(`  ❌ Failed imports: ${failed.length}/${results.length}`);
  console.log(`  📊 Total functions tested: ${totalFunctions}`);
  
  if (failed.length > 0) {
    console.log('\n❌ Failed modules:');
    failed.forEach(f => console.log(`  - ${f.module}: ${f.error}`));
  }
  
  return {
    success: failed.length === 0,
    totalModules: results.length,
    successfulModules: successful.length,
    failedModules: failed.length,
    totalFunctions,
    results
  };
}

// Run the test
testModuleImports().then(result => {
  if (result.success) {
    console.log('\n✅ Integration test PASSED!');
    console.log(`All ${result.totalModules} modules imported successfully with ${result.totalFunctions} functions.`);
    process.exit(0);
  } else {
    console.log('\n❌ Integration test FAILED!');
    process.exit(1);
  }
}).catch(error => {
  console.error('Integration test encountered an error:', error);
  process.exit(1);
});