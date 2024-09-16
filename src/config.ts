import { ExternalConnectors } from '@microsoft/microsoft-graph-types';

export const config = {
  connection: {
    id: 'msgraphdocs',
    name: 'Microsoft Graph documentation',
    description: 'Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it.',
    activitySettings: {
      // URL to item resolves track activity such as sharing external items.
      // The recorded activity is used to improve search relevance.
      urlToItemResolvers: [
        {
          urlMatchInfo: {
            baseUrls: [
              'https://app.contoso.com'
            ],
            urlPattern: '/kb/(?<slug>[^/]+)'
          },
          itemId: '{slug}',
          priority: 1
        } as ExternalConnectors.ItemIdResolver
      ]
    },
    searchSettings: {
      searchResultTemplates: [
        {
          id: '',
          priority: 1,
          layout: {}
        }
      ]
    },
    // https://learn.microsoft.com/graph/connecting-external-content-manage-schema
    schema: {
      baseType: 'microsoft.graph.externalItem',
      // Add properties as needed
      properties: [
        {
          name: 'title',
          type: 'string',
          isQueryable: true,
          isSearchable: true,
          isRetrievable: true,
          labels: [
            'title'
          ]
        },
        {
          name: 'url',
          type: 'string',
          isRetrievable: true,
          labels: [
            'url'
          ]
        },
        {
          name: 'iconUrl',
          type: 'string',
          isRetrievable: true,
          labels: [
            'iconUrl'
          ]
        },
        {
          name: 'description',
          type: 'string',
          isQueryable: true,
          isSearchable: true,
          isRetrievable: true
        },
      ]
    }
  } as ExternalConnectors.ExternalConnection
};