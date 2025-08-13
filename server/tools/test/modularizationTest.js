// Test to verify all tools are properly exported from the modularized structure

import * as tools from '../index.js';

// List of all expected tools based on original monolithic file
const expectedTools = [
  // Email tools
  'listEmailsTool',
  'getEmailTool',
  'sendEmailTool',
  'searchEmailsTool',
  'createDraftTool',
  'replyToEmailTool',
  'replyAllTool',
  'forwardEmailTool',
  'deleteEmailTool',
  'moveEmailTool',
  'markAsReadTool',
  'flagEmailTool',
  'categorizeEmailTool',
  'archiveEmailTool',
  'batchProcessEmailsTool',

  // Calendar tools
  'listEventsTool',
  'createEventTool',
  'getEventTool',
  'updateEventTool',
  'deleteEventTool',
  'createRecurringEventTool',
  'findMeetingTimesTool',
  'checkAvailabilityTool',
  'scheduleOnlineMeetingTool',
  'respondToInviteTool',
  'listCalendarsTool',
  'getCalendarViewTool',
  'getBusyTimesTool',
  'buildRecurrencePatternTool',
  'createRecurrenceHelperTool',
  'validateEventDateTimesTool',
  'checkCalendarPermissionsTool',

  // Folder tools
  'listFoldersTool',
  'createFolderTool',
  'renameFolderTool',
  'getFolderStatsTool',

  // Attachment tools
  'listAttachmentsTool',
  'downloadAttachmentTool',
  'addAttachmentTool',
  'scanAttachmentsTool',

  // Utility tools
  'getRateLimitMetricsTool',
  'resetRateLimitMetricsTool',

  // Common utilities
  'clearStylingCache',
  'clearSignatureCache',
  'getStylingCacheStats',
  'applyUserStyling',
  'stylingCache',
  'signatureCache'
];

// Test function
function testModularization() {
  console.log('🔍 Testing modularized tools structure...\n');
  
  const missingTools = [];
  const foundTools = [];
  
  expectedTools.forEach(toolName => {
    if (tools[toolName]) {
      foundTools.push(toolName);
    } else {
      missingTools.push(toolName);
    }
  });
  
  console.log(`✅ Found ${foundTools.length} tools:`);
  foundTools.forEach(tool => console.log(`  - ${tool}`));
  
  if (missingTools.length > 0) {
    console.log(`\n❌ Missing ${missingTools.length} tools:`);
    missingTools.forEach(tool => console.log(`  - ${tool}`));
  } else {
    console.log('\n🎉 All tools successfully exported!');
  }
  
  // Check for any unexpected exports
  const allExports = Object.keys(tools);
  const unexpectedExports = allExports.filter(exportName => 
    !expectedTools.includes(exportName)
  );
  
  if (unexpectedExports.length > 0) {
    console.log(`\n⚠️  Unexpected exports (${unexpectedExports.length}):`);
    unexpectedExports.forEach(exportName => console.log(`  - ${exportName}`));
  }
  
  console.log(`\n📊 Summary:`);
  console.log(`  - Expected: ${expectedTools.length}`);
  console.log(`  - Found: ${foundTools.length}`);
  console.log(`  - Missing: ${missingTools.length}`);
  console.log(`  - Unexpected: ${unexpectedExports.length}`);
  console.log(`  - Total exports: ${allExports.length}`);
  
  return {
    success: missingTools.length === 0,
    expectedCount: expectedTools.length,
    foundCount: foundTools.length,
    missingCount: missingTools.length,
    unexpectedCount: unexpectedExports.length,
    missingTools,
    unexpectedExports
  };
}

// Run the test
const result = testModularization();

// Exit with appropriate code
if (result.success) {
  console.log('\n✅ Modularization test PASSED!');
  process.exit(0);
} else {
  console.log('\n❌ Modularization test FAILED!');
  process.exit(1);
}