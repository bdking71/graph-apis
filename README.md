# graph-apis

## Summary

This is a proof-of-concept application which pulls data from multiple Exchange Group Calendars and displays in a single calendar in SharePoint.

## Screenshots
![Screenshot](https://bdking71.files.wordpress.com/2022/03/calendar.png "Month View")

![Screenshot](https://bdking71.files.wordpress.com/2022/03/calendar2.png "Day View")

![Screenshot](https://bdking71.files.wordpress.com/2022/03/calendar1.png "Event")

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.14-green.svg "SFPx Version 1.14")

## Prerequisites

> Any special pre-requisites?

## Version history

Version|Comments
--------|---------
20220326-0915 | Added the ability to add and display multiple Outlook Group Calendars. 
20220404-0922 | Adding the ability to add and display Multiple SharePoint Calendars.

## Known Issues

Date|Issues Nbr|Description|Status
--------|---------|---------|---------
20220404|202204040921|Recurring Events in Outlook only shows the first event in the series.|Fixed

### "Sharing is Caring"
And whatever you do, whether in word or deed, do it all in the name of the Lord Jesus, giving thanks to God the Father through him. -- Colossians 3:17 (NIV)


### Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- Allows for displaying of multiple Outlook Calendars 
- Allows for displaying of multiple SharePoint Calendars

## References
- [@pnp/spfx-property-control](https://pnp.github.io/sp-dev-fx-property-controls) - This repository provides developers with a set of reusable property pane controls that can be used in their SharePoint Framework (SPFx) solutions.
- [classnames](https://www.npmjs.com/package/classnames) - A simple JavaScript utility for conditionally joining classNames together.
- [Microsoft Graph TypeScript Types](https://www.npmjs.com/package/@microsoft/microsoft-graph-types) - The Microsoft Graph TypeScript definitions enable editors to provide intellisense on Microsoft Graph objects including users, messages, and groups
- [React Awesome Calendar](https://www.npmjs.com/package/react-awesome-calendar) - react-awesome-calendar is a library that allows you to easily add a calendar to your application. React Awesome Calendar also supports the ability to display events.
- [PNP Transition Guide](https://pnp.github.io/pnpjs/getting-started/) - PnPjs is a collection of fluent libraries for consuming SharePoint, Graph, and Office 365 REST APIs in a type-safe way. You can use it within SharePoint Framework, Nodejs, or any JavaScript project. This an open source initiative and we encourage contributions and constructive feedback from the community.
- [(Stack Exchange) REST API + Expand Recurring Calendar Events?](https://sharepoint.stackexchange.com/questions/23221/rest-api-expand-recurring-calendar-events) - Is there a way using the REST API (through javascript) to expand recurring calendar items? Or is there a helpful client side utility to assist with this?
- [@pnp/graph/calendars](https://pnp.github.io/pnpjs/graph/calendars/) - Information on pnp.js and Graph calendars.
- [sharepoint-events-parser](https://www.npmjs.com/package/sharepoint-events-parser) - Recurring events on a SharePoint calendar are not stored individually; instead, the parent event contains the recurrence information stored as XML. The only other way I've seen to get recurrence data from a calendar list on the client side is to use the Lists.asmx web service. This is not necessarily a bad way to go, but working with the CAML for the query and the XML returned from that web service can be burdensome.
- [axios](https://www.npmjs.com/package/axios) - Promise based HTTP client for the browser and node.js
- [Lorem Ipsum Generators](https://loremipsum.io/ultimate-list-of-lorem-ipsum-generators/) - Think classic lorem ipsum is passé? Give your next project a bit more edge with these funny and unique text generators
