/**
 * ============================================================================
 * 🛡️ REVOPS ENTERPRISE OS — SECURE CLIENT RUNNER (THIN CLIENT LOADER)
 * ============================================================================
 * Copyright (c) 2026 RevOps Technologies. All rights reserved.
 * 
 * NOTICE:
 * This software is proprietary and licensed under commercial enterprise terms.
 * The core logic, calculation algorithms, and AI orchestrator are executed
 * through the encrypted private RevOpsCore Library.
 * 
 * Unauthorized copying, distribution, or reverse engineering is strictly prohibited.
 */

// 1. Initialization and Library Binding
const REV_OPS_CORE = RevOpsCoreLib; // Private Library Script ID: 1x_RevOpsCore_Enterprise_Library_ID

function onOpen(e) {
  REV_OPS_CORE.onOpenHandler(e);
}

function onEditInstallable(e) {
  REV_OPS_CORE.onEditHandler(e);
}

function showEnterpriseNavigator() {
  REV_OPS_CORE.showNavigator();
}

function endPlanningSession() {
  REV_OPS_CORE.endSession();
}

function showMigrationWizardDialog() {
  REV_OPS_CORE.showMigrationWizard();
}

function showAddUserDialog() {
  REV_OPS_CORE.showUserDialog();
}

function showRoleSwitcherDialog() {
  REV_OPS_CORE.showRoleSwitcher();
}

function checkAndDispatchSlaAlerts() {
  return REV_OPS_CORE.dispatchAlerts();
}
