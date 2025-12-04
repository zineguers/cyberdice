"use client";

import { useState, useCallback, useRef } from "react";
import {
  Mail,
  AlertCircle,
  CheckCircle2,
  Search,
  Trash2,
  Download,
  Eye,
  FileText,
  Database,
} from "lucide-react";

export default function CyberDicePage() {
  // Connection state
  const [tenantId, setTenantId] = useState("");
  const [clientId, setClientId] = useState("");
  const [clientSecret, setClientSecret] = useState("");
  const [accessToken, setAccessToken] = useState(null);
  const [isConnected, setIsConnected] = useState(false);
  const [connectionError, setConnectionError] = useState("");
  const [isConnecting, setIsConnecting] = useState(false);

  // Search state
  const [sender, setSender] = useState("");
  const [receiver, setReceiver] = useState("");
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [includeAttachments, setIncludeAttachments] = useState(false);
  const [isSearching, setIsSearching] = useState(false);
  const [searchProgress, setSearchProgress] = useState("");

  // Results state
  const [results, setResults] = useState([]);
  const [selectedMessages, setSelectedMessages] = useState(new Set());
  const [previewMessage, setPreviewMessage] = useState(null);

  // Delete state
  const [purgeText, setPurgeText] = useState("");
  const [isDeleting, setIsDeleting] = useState(false);
  const [showPurgeConfirm, setShowPurgeConfirm] = useState(false);

  // Connect to Microsoft Graph
  const handleConnect = useCallback(async () => {
    setConnectionError("");
    setIsConnecting(true);

    try {
      const response = await fetch("/api/graph/auth", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tenantId, clientId, clientSecret }),
      });

      const data = await response.json();

      if (!response.ok) {
        setConnectionError(data.error || "Authentication failed");
        setIsConnected(false);
        setAccessToken(null);
        return;
      }

      setAccessToken(data.accessToken);
      setIsConnected(true);
      setConnectionError("");
    } catch (error) {
      setConnectionError(error.message);
      setIsConnected(false);
    } finally {
      setIsConnecting(false);
    }
  }, [tenantId, clientId, clientSecret]);

  // Search emails
  const handleSearch = useCallback(async () => {
    setIsSearching(true);
    setSearchProgress("Initializing search...");
    setResults([]);
    setSelectedMessages(new Set());

    try {
      const response = await fetch("/api/graph/search", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          accessToken,
          sender,
          receiver,
          subject,
          body,
          startDate,
          endDate,
          includeAttachments,
        }),
      });

      const data = await response.json();

      if (!response.ok) {
        alert(data.error || "Search failed");
        return;
      }

      setResults(data.results || []);
      setSearchProgress("");
    } catch (error) {
      alert(error.message);
    } finally {
      setIsSearching(false);
    }
  }, [
    accessToken,
    sender,
    receiver,
    subject,
    body,
    startDate,
    endDate,
    includeAttachments,
  ]);

  // Toggle message selection
  const toggleSelect = useCallback((messageId) => {
    setSelectedMessages((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(messageId)) {
        newSet.delete(messageId);
      } else {
        newSet.add(messageId);
      }
      return newSet;
    });
  }, []);

  // Select all
  const handleSelectAll = useCallback(() => {
    if (selectedMessages.size === results.length) {
      setSelectedMessages(new Set());
    } else {
      setSelectedMessages(new Set(results.map((r) => r.id)));
    }
  }, [results, selectedMessages]);

  // Preview selected
  const handlePreview = useCallback(async () => {
    if (selectedMessages.size === 0) return;

    const messageId = Array.from(selectedMessages)[0];
    const message = results.find((r) => r.id === messageId);

    if (!message) return;

    try {
      const response = await fetch("/api/graph/message-details", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          accessToken,
          mailbox: message.mailbox,
          messageId: message.id,
        }),
      });

      const data = await response.json();
      if (response.ok) {
        setPreviewMessage(data);
      }
    } catch (error) {
      alert("Failed to load preview");
    }
  }, [selectedMessages, results, accessToken]);

  // Download as EML
  const handleDownloadEML = useCallback(async () => {
    if (selectedMessages.size === 0) return;

    for (const messageId of selectedMessages) {
      const message = results.find((r) => r.id === messageId);
      if (!message) continue;

      try {
        const response = await fetch("/api/graph/download-eml", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            accessToken,
            mailbox: message.mailbox,
            messageId: message.id,
          }),
        });

        if (response.ok) {
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = `${message.subject || "message"}.eml`;
          a.click();
          window.URL.revokeObjectURL(url);
        }
      } catch (error) {
        console.error("Download failed:", error);
      }
    }
  }, [selectedMessages, results, accessToken]);

  // Download attachments
  const handleDownloadAttachments = useCallback(async () => {
    if (selectedMessages.size === 0) return;

    for (const messageId of selectedMessages) {
      const message = results.find((r) => r.id === messageId);
      if (!message || !message.hasAttachments) continue;

      try {
        const response = await fetch("/api/graph/attachments", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            accessToken,
            mailbox: message.mailbox,
            messageId: message.id,
          }),
        });

        const data = await response.json();
        if (response.ok && data.attachments) {
          data.attachments.forEach((att) => {
            if (att.contentBytes) {
              const blob = new Blob([
                Uint8Array.from(atob(att.contentBytes), (c) => c.charCodeAt(0)),
              ]);
              const url = window.URL.createObjectURL(blob);
              const a = document.createElement("a");
              a.href = url;
              a.download = att.name;
              a.click();
              window.URL.revokeObjectURL(url);
            }
          });
        }
      } catch (error) {
        console.error("Download attachments failed:", error);
      }
    }
  }, [selectedMessages, results, accessToken]);

  // Soft delete
  const handleSoftDelete = useCallback(async () => {
    if (selectedMessages.size === 0) return;
    if (!confirm(`Move ${selectedMessages.size} email(s) to Deleted Items?`))
      return;

    setIsDeleting(true);
    const messagesToDelete = results.filter((r) => selectedMessages.has(r.id));

    try {
      const response = await fetch("/api/graph/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          accessToken,
          messages: messagesToDelete,
          permanent: false,
        }),
      });

      const data = await response.json();
      if (response.ok) {
        setResults((prev) => prev.filter((r) => !selectedMessages.has(r.id)));
        setSelectedMessages(new Set());
        alert(`Successfully deleted ${data.deleted} email(s)`);
      }
    } catch (error) {
      alert("Delete failed: " + error.message);
    } finally {
      setIsDeleting(false);
    }
  }, [selectedMessages, results, accessToken]);

  // Hard delete / purge
  const handleHardDelete = useCallback(async () => {
    if (selectedMessages.size === 0) return;

    setIsDeleting(true);
    const messagesToDelete = results.filter((r) => selectedMessages.has(r.id));

    try {
      const response = await fetch("/api/graph/delete", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          accessToken,
          messages: messagesToDelete,
          permanent: true,
        }),
      });

      const data = await response.json();
      if (response.ok) {
        setResults((prev) => prev.filter((r) => !selectedMessages.has(r.id)));
        setSelectedMessages(new Set());
        setPurgeText("");
        setShowPurgeConfirm(false);
        alert(`Successfully purged ${data.deleted} email(s)`);
      }
    } catch (error) {
      alert("Purge failed: " + error.message);
    } finally {
      setIsDeleting(false);
    }
  }, [selectedMessages, results, accessToken]);

  // Export to CSV
  const handleExportCSV = useCallback(() => {
    if (results.length === 0) return;

    const csvRows = [
      ["Date", "From", "To", "Subject", "Has Attachments", "Mailbox"],
    ];

    results.forEach((msg) => {
      csvRows.push([
        new Date(msg.receivedDateTime).toLocaleString(),
        msg.from?.emailAddress?.address || "",
        msg.toRecipients?.map((t) => t.emailAddress.address).join("; ") || "",
        msg.subject || "",
        msg.hasAttachments ? "Yes" : "No",
        msg.mailbox || "",
      ]);
    });

    const csvContent = csvRows
      .map((row) => row.map((cell) => `"${cell}"`).join(","))
      .join("\n");

    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "email-results.csv";
    a.click();
    window.URL.revokeObjectURL(url);
  }, [results]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100">
      {/* Top Bar */}
      <div className="w-full bg-gradient-to-r from-[#ff3131] to-[#ff5555] py-4 px-6 shadow-lg">
        <h1 className="text-white text-2xl font-bold tracking-wide">
          CYBER DICE
        </h1>
      </div>

      <div className="max-w-[1600px] mx-auto p-6 space-y-6">
        {/* Connection Card */}
        <div className="bg-white rounded-lg shadow-md overflow-hidden">
          <div className="bg-gradient-to-r from-purple-600 to-purple-700 px-6 py-4">
            <h2 className="text-white text-lg font-semibold">
              Microsoft Graph Connection
            </h2>
          </div>
          <div className="p-6 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Tenant ID
                </label>
                <input
                  type="text"
                  value={tenantId}
                  onChange={(e) => setTenantId(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                  placeholder="Enter Tenant ID"
                  disabled={isConnected}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Client ID
                </label>
                <input
                  type="text"
                  value={clientId}
                  onChange={(e) => setClientId(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                  placeholder="Enter Client ID"
                  disabled={isConnected}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Client Secret
                </label>
                <input
                  type="password"
                  value={clientSecret}
                  onChange={(e) => setClientSecret(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                  placeholder="Enter Client Secret"
                  disabled={isConnected}
                />
              </div>
            </div>

            <div className="flex items-center gap-4">
              <button
                onClick={handleConnect}
                disabled={
                  isConnecting ||
                  isConnected ||
                  !tenantId ||
                  !clientId ||
                  !clientSecret
                }
                className="px-6 py-2 bg-gray-700 text-white rounded-lg hover:bg-gray-800 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
              >
                {isConnecting
                  ? "Connecting..."
                  : isConnected
                    ? "Connected"
                    : "Connect to Microsoft Graph"}
              </button>

              <div className="flex items-center gap-2">
                {isConnected ? (
                  <>
                    <CheckCircle2 className="w-5 h-5 text-green-500" />
                    <span className="text-green-600 font-medium">
                      Connected
                    </span>
                  </>
                ) : (
                  <>
                    <AlertCircle className="w-5 h-5 text-red-500" />
                    <span className="text-red-600 font-medium">
                      Disconnected
                    </span>
                  </>
                )}
              </div>
            </div>

            {connectionError && (
              <div className="p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
                {connectionError}
              </div>
            )}
          </div>
        </div>

        {/* Search Card */}
        <div className="bg-white rounded-lg shadow-md overflow-hidden">
          <div className="bg-gradient-to-r from-green-600 to-green-700 px-6 py-4">
            <h2 className="text-white text-lg font-semibold">
              Search Criteria
            </h2>
          </div>
          <div className="p-6 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Sender
                </label>
                <input
                  type="text"
                  value={sender}
                  onChange={(e) => setSender(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                  placeholder="sender@example.com"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Receiver (comma-separated)
                </label>
                <input
                  type="text"
                  value={receiver}
                  onChange={(e) => setReceiver(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                  placeholder="user1@example.com, user2@example.com"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Subject Contains
                </label>
                <input
                  type="text"
                  value={subject}
                  onChange={(e) => setSubject(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                  placeholder="Enter subject keyword"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Body Contains
                </label>
                <input
                  type="text"
                  value={body}
                  onChange={(e) => setBody(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                  placeholder="Enter body keyword"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Start Date
                </label>
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  End Date
                </label>
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent"
                />
              </div>
            </div>

            <div className="flex items-center gap-2">
              <input
                type="checkbox"
                id="attachments"
                checked={includeAttachments}
                onChange={(e) => setIncludeAttachments(e.target.checked)}
                className="w-4 h-4 text-green-600 border-gray-300 rounded focus:ring-green-500"
              />
              <label
                htmlFor="attachments"
                className="text-sm font-medium text-gray-700"
              >
                Only show emails with attachments
              </label>
            </div>

            <div className="flex items-center gap-4">
              <button
                onClick={handleSearch}
                disabled={!isConnected || isSearching}
                className="px-6 py-2 bg-gray-700 text-white rounded-lg hover:bg-gray-800 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center gap-2"
              >
                <Search className="w-4 h-4" />
                {isSearching ? "Searching..." : "Search Emails"}
              </button>

              {searchProgress && (
                <span className="text-sm text-gray-600">{searchProgress}</span>
              )}
            </div>
          </div>
        </div>

        {/* Results Card */}
        <div className="bg-white rounded-lg shadow-md overflow-hidden">
          <div className="bg-gradient-to-r from-gray-800 to-gray-900 px-6 py-4">
            <h2 className="text-white text-lg font-semibold">Search Results</h2>
          </div>

          {/* Actions Panel - ON TOP */}
          {results.length > 0 && (
            <div className="bg-gradient-to-r from-purple-50 to-purple-100 border-b border-purple-200 p-4">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <div className="bg-purple-600 text-white px-4 py-1 rounded-full text-sm font-medium">
                    {selectedMessages.size} email(s) selected
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3">
                <button
                  onClick={handlePreview}
                  disabled={selectedMessages.size === 0}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <Eye className="w-4 h-4" />
                  Preview
                </button>

                <button
                  onClick={handleDownloadEML}
                  disabled={selectedMessages.size === 0}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <Download className="w-4 h-4" />
                  Download EML
                </button>

                <button
                  onClick={handleDownloadAttachments}
                  disabled={selectedMessages.size === 0}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <FileText className="w-4 h-4" />
                  Attachments
                </button>

                <button
                  onClick={handleSoftDelete}
                  disabled={selectedMessages.size === 0 || isDeleting}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <Trash2 className="w-4 h-4" />
                  Soft Delete
                </button>

                <button
                  onClick={handleExportCSV}
                  disabled={results.length === 0}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm"
                >
                  <Database className="w-4 h-4" />
                  Export CSV
                </button>

                <div className="col-span-2 flex flex-col gap-2">
                  <input
                    type="text"
                    value={purgeText}
                    onChange={(e) => setPurgeText(e.target.value)}
                    placeholder="Type 'purge' to enable"
                    className="px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-red-500 focus:border-transparent"
                  />
                  <button
                    onClick={() => {
                      if (
                        confirm(
                          `PERMANENTLY DELETE ${selectedMessages.size} email(s)? This cannot be undone!`,
                        )
                      ) {
                        handleHardDelete();
                      }
                    }}
                    disabled={
                      selectedMessages.size === 0 ||
                      purgeText !== "purge" ||
                      isDeleting
                    }
                    className="px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center gap-2 text-sm font-medium"
                  >
                    <Trash2 className="w-4 h-4" />
                    PURGE
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Results Table */}
          <div className="p-6">
            {results.length === 0 ? (
              <div className="text-center py-12">
                <Mail className="w-16 h-16 text-gray-300 mx-auto mb-4" />
                <p className="text-gray-500 text-lg">No messages found</p>
                <p className="text-gray-400 text-sm mt-2">
                  Use the search criteria above to find emails
                </p>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead>
                    <tr className="border-b border-gray-200">
                      <th className="text-left p-3">
                        <input
                          type="checkbox"
                          checked={
                            selectedMessages.size === results.length &&
                            results.length > 0
                          }
                          onChange={handleSelectAll}
                          className="w-4 h-4 text-purple-600 border-gray-300 rounded focus:ring-purple-500"
                        />
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        Date
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        From
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        To
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        Subject
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        Attachments
                      </th>
                      <th className="text-left p-3 text-sm font-semibold text-gray-700">
                        Mailbox
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.map((msg) => (
                      <tr
                        key={msg.id}
                        className="border-b border-gray-100 hover:bg-gray-50"
                      >
                        <td className="p-3">
                          <input
                            type="checkbox"
                            checked={selectedMessages.has(msg.id)}
                            onChange={() => toggleSelect(msg.id)}
                            className="w-4 h-4 text-purple-600 border-gray-300 rounded focus:ring-purple-500"
                          />
                        </td>
                        <td className="p-3 text-sm text-gray-600">
                          {new Date(msg.receivedDateTime).toLocaleString()}
                        </td>
                        <td className="p-3 text-sm text-gray-900">
                          {msg.from?.emailAddress?.address || "N/A"}
                        </td>
                        <td className="p-3 text-sm text-gray-900">
                          {msg.toRecipients
                            ?.map((t) => t.emailAddress.address)
                            .join(", ") || "N/A"}
                        </td>
                        <td className="p-3 text-sm text-gray-900 max-w-md truncate">
                          {msg.subject || "(No Subject)"}
                        </td>
                        <td className="p-3 text-sm text-gray-600">
                          {msg.hasAttachments ? "✓" : "—"}
                        </td>
                        <td className="p-3 text-sm text-gray-600">
                          {msg.mailbox}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Preview Modal */}
      {previewMessage && (
        <div
          className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50"
          onClick={() => setPreviewMessage(null)}
        >
          <div
            className="bg-white rounded-lg max-w-4xl w-full max-h-[80vh] overflow-hidden"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="bg-gradient-to-r from-gray-800 to-gray-900 px-6 py-4 flex justify-between items-center">
              <h3 className="text-white text-lg font-semibold">
                Email Preview
              </h3>
              <button
                onClick={() => setPreviewMessage(null)}
                className="text-white hover:text-gray-300"
              >
                <span className="text-2xl">×</span>
              </button>
            </div>
            <div className="p-6 overflow-y-auto max-h-[calc(80vh-80px)]">
              <div className="space-y-3">
                <div>
                  <span className="font-semibold text-gray-700">Subject:</span>
                  <p className="text-gray-900">
                    {previewMessage.message?.subject || "(No Subject)"}
                  </p>
                </div>
                <div>
                  <span className="font-semibold text-gray-700">From:</span>
                  <p className="text-gray-900">
                    {previewMessage.message?.from?.emailAddress?.address}
                  </p>
                </div>
                <div>
                  <span className="font-semibold text-gray-700">To:</span>
                  <p className="text-gray-900">
                    {previewMessage.message?.toRecipients
                      ?.map((t) => t.emailAddress.address)
                      .join(", ")}
                  </p>
                </div>
                <div>
                  <span className="font-semibold text-gray-700">Date:</span>
                  <p className="text-gray-900">
                    {new Date(
                      previewMessage.message?.receivedDateTime,
                    ).toLocaleString()}
                  </p>
                </div>
                {previewMessage.attachments?.length > 0 && (
                  <div>
                    <span className="font-semibold text-gray-700">
                      Attachments:
                    </span>
                    <ul className="list-disc list-inside text-gray-900">
                      {previewMessage.attachments.map((att, i) => (
                        <li key={i}>{att.name}</li>
                      ))}
                    </ul>
                  </div>
                )}
                <div>
                  <span className="font-semibold text-gray-700">Body:</span>
                  <div className="mt-2 p-4 bg-gray-50 rounded border border-gray-200">
                    {previewMessage.message?.body?.contentType === "html" ? (
                      <div
                        dangerouslySetInnerHTML={{
                          __html: previewMessage.message.body.content,
                        }}
                      />
                    ) : (
                      <p className="text-gray-900 whitespace-pre-wrap">
                        {previewMessage.message?.body?.content ||
                          previewMessage.message?.bodyPreview}
                      </p>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
