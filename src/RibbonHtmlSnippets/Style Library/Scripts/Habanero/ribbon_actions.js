/**
 * Module for adding a ribbon action allowing users to inject HTML snippets when authoring fields AND content editor web parts. Designed to address the limitation of the Reusable Content
 * ribbon function, which doesn't work with content editor web parts.
 *
 * Author: Mark Bice
 */

/*! Habanero - Licensed under MIT */


/**
* Ribbon Actions module
* @namespace Hcf.RibbonActions
*/
(function (window, document, undefined) {
	'use strict';

	var snippetsData;


	//#region $PRIVATE METHODS

	/**
	* Initialization method
	*/
	function init() {
		// Once SP and our ribbon libraries have loaded then add the HTML snippet insertion button to the ribbon
		SP.SOD.executeOrDelayUntilScriptLoaded(function () {
			SP.SOD.executeOrDelayUntilScriptLoaded(function () {
				getReusableContent(function (data) {
					if (data.length > 0) {
						snippetsData = data;
						addInsertHtmlSnippetAction();
					}
				});
			}, 'cui.js');
		}, 'sp.js');
	}


	/**
	* Fetches reusable content list using CSOM. Using this instead of the REST API to make it easier to use non-jQuery code and to make it backwards compatible with SP 2010, which has some issues with
	* publishing HTML fields and the REST API.
	*
	* @param {object} callback Function to call back once list items successfully retrieved
	*/
	function getReusableContent(callback) {
		var ctx = new SP.ClientContext.get_current(),
			web = ctx.get_web(),
			oList = web.get_lists().getByTitle('Reusable Content'),
			oListItems,
			oListItem,
			caml = new SP.CamlQuery(),
			results = [];

		caml.set_viewXml('');
		oListItems = oList.getItems(caml);

		ctx.load(oListItems);

		ctx.executeQueryAsync(
			function (sender, args) {
				var listItemEnumerator = oListItems.getEnumerator();
				while (listItemEnumerator.moveNext()) {
					oListItem = listItemEnumerator.get_current();
					results.push({
						Title: oListItem.get_item('Title'),
						ReusableHtml: oListItem.get_item('ReusableHtml')

					});
				}

				callback(results);  // Send results back to our callback
			},
			function (sender, args) {
				console.log('Error adding ribbon action for HTML snippets:' + args.get_message());
			}
		);
	}


	/**
	* Creates our ribbon action component
	*/
	function addInsertHtmlSnippetAction() {
		Type.registerNamespace('Hcf.RibbonActions.InsertHtmlSnippet.PageComponent');  // Register a namespace for this component

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent = function () {
			Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.initializeBase(this);
		};

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.initialize = function () {
			SP.SOD.executeOrDelayUntilScriptLoaded(Function.createDelegate(null, Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.initializePageComponent), 'SP.Ribbon.js');
		};

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.initializePageComponent = function () {
			var ribbonPageManager = SP.Ribbon.PageManager.get_instance();
			if (null !== ribbonPageManager) {
				ribbonPageManager.addPageComponent(Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.instance);
			}
		};

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.prototype = {
			init: function () { },
			getFocusedCommands: function () {
				return [];
			},
			getGlobalCommands: function () {
				return ['Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet', 'Hcf.RibbonActions.InsertHtmlSnippet.LoadSnippets'];
			},
			canHandleCommand: function (commandId) {
				if (commandId === 'Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet' ||
					commandId === 'Hcf.RibbonActions.InsertHtmlSnippet.LoadSnippets') {
					return true;
				}
				else {
					return false;
				}
			},
			handleCommand: function (commandId, properties, sequence) {
				if (commandId === 'Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet') {
					Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet(properties['CommandValueId']);
				}
				else if (commandId === 'Hcf.RibbonActions.InsertHtmlSnippet.LoadSnippets') {
					properties.PopulationXML = Hcf.RibbonActions.InsertHtmlSnippet.GetAvailableSnippets();
				}
			},
			isFocusable: function () {
				return true;
			},
			receiveFocus: function () {
				return true;
			},
			yieldFocus: function () {
				return true;
			}
		};

		Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet = function (snippet) {
			RTE.Cursor.get_range().replaceHtml(snippet);
		};

		Hcf.RibbonActions.InsertHtmlSnippet.GetAvailableSnippets = function () {
			var sb = new Sys.StringBuilder(),
				item;

			// If we have snippets to show then construct the flyout options (one for each snippet)
			if (snippetsData.length > 0) {
				sb.append('<Menu Id=\'Ribbon.EditingTools.CPInsert.Content.InsertHtmlSnippet.Menu\'>');
				sb.append('<MenuSection Id=\'Ribbon.EditingTools.CPInsert.Content.InsertHtmlSnippet.Menu.Symbols\'>');
				sb.append('<Controls Id=\'Ribbon.EditingTools.CPInsert.Content.InsertHtmlSnippet.Menu.Symbols.Controls\'>');

				for (var i = 0; i < snippetsData.length; i++) {
					item = snippetsData[i];
					sb.append('<Button Id=\'Ribbon.EditingTools.CPInsert.Content.InsertSymbol.Menu.Bookmarks' + i + '\' LabelText=\'' + item.Title + '\' Command=\'Hcf.RibbonActions.InsertHtmlSnippet.InsertSnippet\' CommandType=\'General\' CommandValueId=\'' + htmlEscape(item.ReusableHtml) + '\' />');
				}

				sb.append('</Controls>');
				sb.append('</MenuSection>');
				sb.append('</Menu>');
			}

			return sb.toString();
		};

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.registerClass('Hcf.RibbonActions.InsertHtmlSnippet.PageComponent', CUI.Page.PageComponent);
		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.instance = new Hcf.RibbonActions.InsertHtmlSnippet.PageComponent();

		Hcf.RibbonActions.InsertHtmlSnippet.PageComponent.initialize();
	}

	//#endregion



	//#region UTILITY

	/**
	* Escapes HTML strings
	* @param {string} str Html string to escape
	*/
	function htmlEscape(str) {
		return String(str)
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
	}

	//#endregion


	// run our initialization method
	init();

})(Hcf, window, document);