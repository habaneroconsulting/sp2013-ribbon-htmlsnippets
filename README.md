#Ribbon action for inserting HTML snippets (SharePoint 2013 on premise and O365)

This is a sample sandbox solution that adds a custom ribbon button allowing authors to insert HTML snippets into both fields AND content editor web parts. This addresses the out of the box limitation of SharePoint's Reusable Content feature, which only works with fields, not content editors. Although this won't have any of the linking capabilities of Reusable Content, it's purpose is more of an HTML injector where a copy is pasted into the selected region (no linking). This is a sandbox WSP and works with both on premise and Office 365. Note that Reusable Content is a publishing concept so if you are working within a team site the Reusable Content list will not be there by default. You can create the list manually and this feature will still add the button to be used on team sites.

##Getting Started

Pull all the files from the "src" folder and open up RibbonHtmlSnippets.sln using Visual Studio (built using VS 2012). There is a project called RibbonHtmlSnippets. Simply build, point at your site collection and then deploy it. 

Once deployed, edit a page and highlight either a field or the content area of a content editor web part. On the ribbon go to the Embed tab. You'll see there is a ribbon action named "Insert Snippet". This pulls from the same Reusable Content library by default so you'll see the same options in the flyout for this button as you do in the Reusable Content flyout. Choose one of the options in the flyout and the snippet will be pasted into the field/content editor you are highlighted on. The solution is hardcoded to pull from the Reusable Content library relative to your current site. If you want to pull from a different location or always pull from the root site collection you'll have to modify the "getReusableContent()" function to do so.

##Deploying to O365

1. You can either manually upload the generated WSP to your O365 tenant or use Visual Studio's "Publish" command from the right click menu.
2. Once published go to the Solutions page in your tenant and make sure the solution is activated.
3. To re-publish from VS the solution needs to be deleted entirely from your SharePoint tenant.

## Support

If you have a bug, or a feature request, please post in the [issue tracker](https://github.com/habaneroconsulting/sp2013-ribbon-htmlsnippets/issues).

## License

Copyright (c) 2013 Habanero Consulting Group

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: 

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.