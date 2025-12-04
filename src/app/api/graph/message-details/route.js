export async function POST(request) {
  try {
    const { accessToken, mailbox, messageId } = await request.json();

    if (!accessToken || !mailbox || !messageId) {
      return Response.json({ error: "Invalid request" }, { status: 400 });
    }

    // Get full message details
    const messageUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${messageId}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,bodyPreview`;

    const response = await fetch(messageUrl, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });

    if (!response.ok) {
      return Response.json(
        { error: "Failed to fetch message" },
        { status: response.status },
      );
    }

    const message = await response.json();

    // Get attachments if any
    let attachments = [];
    if (message.hasAttachments) {
      const attachmentsUrl = `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${messageId}/attachments`;
      const attachmentsResponse = await fetch(attachmentsUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (attachmentsResponse.ok) {
        const attachmentsData = await attachmentsResponse.json();
        attachments = attachmentsData.value || [];
      }
    }

    return Response.json({
      success: true,
      message,
      attachments,
    });
  } catch (error) {
    console.error("Message details error:", error);
    return Response.json(
      { error: error.message || "Failed to fetch message details" },
      { status: 500 },
    );
  }
}
