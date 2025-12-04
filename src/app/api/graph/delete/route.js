export async function POST(request) {
  try {
    const { accessToken, messages, permanent } = await request.json();

    if (!accessToken || !messages || !Array.isArray(messages)) {
      return Response.json({ error: "Invalid request" }, { status: 400 });
    }

    const results = [];
    const errors = [];

    for (const message of messages) {
      try {
        const deleteUrl = `https://graph.microsoft.com/v1.0/users/${message.mailbox}/messages/${message.id}`;

        const response = await fetch(deleteUrl, {
          method: "DELETE",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
        });

        if (response.ok || response.status === 204) {
          results.push({ id: message.id, success: true });
        } else {
          const error = await response.text();
          errors.push({ id: message.id, error });
        }
      } catch (error) {
        errors.push({ id: message.id, error: error.message });
      }
    }

    return Response.json({
      success: true,
      deleted: results.length,
      errors: errors,
    });
  } catch (error) {
    console.error("Delete error:", error);
    return Response.json(
      { error: error.message || "Delete failed" },
      { status: 500 },
    );
  }
}
