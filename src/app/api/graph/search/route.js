export async function POST(request) {
  try {
    const {
      accessToken,
      sender,
      receiver,
      subject,
      body,
      startDate,
      endDate,
      includeAttachments,
    } = await request.json();

    if (!accessToken) {
      return Response.json({ error: "Access token required" }, { status: 401 });
    }

    let mailboxes = [];

    // Get all mailboxes or specific ones
    if (receiver && receiver.trim()) {
      mailboxes = receiver
        .split(",")
        .map((email) => email.trim())
        .filter(Boolean);
    } else {
      // Get all users in tenant
      const usersResponse = await fetch(
        "https://graph.microsoft.com/v1.0/users?$select=mail,userPrincipalName&$top=999",
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        },
      );

      if (!usersResponse.ok) {
        const error = await usersResponse.json();
        return Response.json(
          { error: error.error?.message || "Failed to fetch users" },
          { status: usersResponse.status },
        );
      }

      const usersData = await usersResponse.json();
      mailboxes = usersData.value
        .filter((user) => user.mail)
        .map((user) => user.mail);
    }

    const allResults = [];
    const errors = [];

    // Search each mailbox
    for (let i = 0; i < mailboxes.length; i++) {
      const mailbox = mailboxes[i];

      try {
        // Build filter query
        const filters = [];

        if (startDate) {
          filters.push(`receivedDateTime ge ${startDate}T00:00:00Z`);
        }
        if (endDate) {
          filters.push(`receivedDateTime le ${endDate}T23:59:59Z`);
        }

        let filterString =
          filters.length > 0 ? `$filter=${filters.join(" and ")}` : "";

        // Build search query for subject and body
        const searchTerms = [];
        if (subject) {
          searchTerms.push(`subject:"${subject}"`);
        }
        if (body) {
          searchTerms.push(`body:"${body}"`);
        }
        if (sender) {
          searchTerms.push(`from:"${sender}"`);
        }

        const searchString =
          searchTerms.length > 0
            ? `$search="${searchTerms.join(" AND ")}"`
            : "";

        // Combine filter and search
        const queryParams = [
          filterString,
          searchString,
          "$select=id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,body",
          "$top=999",
        ]
          .filter(Boolean)
          .join("&");

        const messagesUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/messages?${queryParams}`;

        let hasMore = true;
        let nextLink = messagesUrl;

        while (hasMore) {
          const messagesResponse = await fetch(nextLink, {
            headers: {
              Authorization: `Bearer ${accessToken}`,
              "Content-Type": "application/json",
            },
          });

          if (!messagesResponse.ok) {
            errors.push({ mailbox, error: "Failed to fetch messages" });
            break;
          }

          const messagesData = await messagesResponse.json();

          // Filter by attachments if needed
          let messages = messagesData.value || [];
          if (includeAttachments) {
            messages = messages.filter((msg) => msg.hasAttachments);
          }

          // Add mailbox info to each message
          messages.forEach((msg) => {
            allResults.push({
              ...msg,
              mailbox: mailbox,
            });
          });

          // Check for pagination
          if (messagesData["@odata.nextLink"]) {
            nextLink = messagesData["@odata.nextLink"];
          } else {
            hasMore = false;
          }
        }
      } catch (error) {
        errors.push({ mailbox, error: error.message });
      }
    }

    return Response.json({
      success: true,
      results: allResults,
      totalMailboxes: mailboxes.length,
      errors: errors,
    });
  } catch (error) {
    console.error("Search error:", error);
    return Response.json(
      { error: error.message || "Search failed" },
      { status: 500 },
    );
  }
}
