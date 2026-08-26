/**
 * @fileoverview Helpers de HtmlService compartilhados
 * @version 1.0.0
 *
 * include(): padrão oficial GAS para compartilhar CSS/JS entre sidebars.
 * Requer que a sidebar seja aberta via createTemplateFromFile(...).evaluate().
 *
 * USO no HTML:
 *   <?!= include('SharedStyles'); ?>
 *   <?!= include('SharedScripts'); ?>
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
