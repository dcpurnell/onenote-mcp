/**
 * Test fixtures for Microsoft Graph API responses
 */

export const mockNotebooks = {
  value: [
    {
      id: 'notebook-1',
      displayName: 'Personal Notebook',
      createdDateTime: '2024-01-15T10:00:00Z',
      lastModifiedDateTime: '2024-02-20T15:30:00Z',
      links: {
        oneNoteWebUrl: {
          href: 'https://onenote.com/notebook1'
        }
      }
    },
    {
      id: 'notebook-2',
      displayName: 'Work Notes',
      createdDateTime: '2024-02-01T09:00:00Z',
      lastModifiedDateTime: '2024-02-22T12:00:00Z',
      links: {
        oneNoteWebUrl: {
          href: 'https://onenote.com/notebook2'
        }
      }
    }
  ]
};

export const mockSections = {
  value: [
    {
      id: 'section-1',
      displayName: 'Quick Notes',
      createdDateTime: '2024-01-15T11:00:00Z',
      lastModifiedDateTime: '2024-02-20T16:00:00Z',
      pagesUrl: 'https://graph.microsoft.com/v1.0/me/onenote/sections/section-1/pages'
    },
    {
      id: 'section-2',
      displayName: 'Meeting Notes',
      createdDateTime: '2024-02-01T10:00:00Z',
      lastModifiedDateTime: '2024-02-22T14:00:00Z',
      pagesUrl: 'https://graph.microsoft.com/v1.0/me/onenote/sections/section-2/pages'
    }
  ]
};

export const mockPages = {
  value: [
    {
      id: 'page-1',
      title: 'Daily Note - 2/22/26',
      createdDateTime: '2024-02-22T08:00:00Z',
      lastModifiedDateTime: '2024-02-22T09:30:00Z',
      contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-1/content',
      links: {
        oneNoteWebUrl: {
          href: 'https://onenote.com/page1'
        }
      },
      createdByAppId: 'app-123',
      lastModifiedByAppId: 'app-123'
    },
    {
      id: 'page-2',
      title: 'Project Planning',
      createdDateTime: '2024-02-21T14:00:00Z',
      lastModifiedDateTime: '2024-02-22T10:00:00Z',
      contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-2/content',
      links: {
        oneNoteWebUrl: {
          href: 'https://onenote.com/page2'
        }
      },
      createdByAppId: 'app-456',
      lastModifiedByAppId: 'app-456'
    }
  ]
};

export const mockPageContentHTML = `
<!DOCTYPE html>
<html>
<head>
  <title>Daily Note - 2/22/26</title>
</head>
<body>
  <div>
    <h1>Daily Note - 2/22/26</h1>
    <p>Today's tasks:</p>
    <ul>
      <li>Review pull requests</li>
      <li>Team meeting at 2pm</li>
      <li>Update documentation</li>
    </ul>
    <h2>Notes from standup</h2>
    <p>Discussed the new OneNote MCP server implementation.</p>
  </div>
</body>
</html>
`;

export const mockPaginatedResponse = {
  '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/onenote/pages?$skip=10',
  value: mockPages.value
};

export const mockUserInfo = {
  displayName: 'Doug Purnell',
  mail: 'dpurnell@elon.edu',
  userPrincipalName: 'dpurnell@elon.edu',
  id: 'user-123'
};

export const mockTokenData = {
  token: 'mock-access-token-12345',
  clientId: '14d82eec-204b-4c2f-b7e8-296a70dab67e',
  scopes: ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'],
  createdAt: '2026-02-22T17:35:32.004Z',
  expiresOn: '2026-02-22T18:48:22.000Z'
};

export const mockErrorResponses = {
  unauthorized: {
    error: {
      code: 'Unauthorized',
      message: 'Access token is missing or invalid',
      innerError: {
        'request-id': 'req-123',
        date: '2024-02-22T10:00:00Z'
      }
    }
  },
  notFound: {
    error: {
      code: 'ResourceNotFound',
      message: 'The specified resource was not found',
      innerError: {
        'request-id': 'req-456',
        date: '2024-02-22T10:00:00Z'
      }
    }
  },
  rateLimit: {
    error: {
      code: 'TooManyRequests',
      message: 'Rate limit exceeded',
      innerError: {
        'request-id': 'req-789',
        date: '2024-02-22T10:00:00Z'
      }
    }
  }
};
