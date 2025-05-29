import * as React from "react";
import { useEffect, useState } from "react";
import { PrimaryButton, TextField } from "@fluentui/react";

interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = ({ title }) => {
  const [senderName, setSenderName] = useState<string>("(Loading sender)");
  const [subject, setSubject] = useState<string>("(Loading subject)");
  const [buildVersion, setBuildVersion] = useState<string>("(Fetching...)");

  useEffect(() => {
    Office.onReady(() => {
      const item = Office.context?.mailbox?.item;

      console.log("Office.context.mailbox.item:", item);

      if (!item) {
        console.warn("No item found in Office.context.mailbox.");
        return;
      }

      // Compose mode
      if (typeof item.subject?.getAsync === "function") {
        console.log("Compose mode detected (getAsync found).");

        item.subject.getAsync((subjectResult) => {
          if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Subject (compose):", subjectResult.value);
            setSubject(subjectResult.value || "(No Subject)");
          } else {
            console.error("Failed to get subject:", subjectResult.error);
            setSubject("(Error reading subject)");
          }
        });

        if (typeof item.from?.getAsync === "function") {
          item.from.getAsync((fromResult) => {
            if (fromResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Sender (compose):", fromResult.value);
              setSenderName(fromResult.value?.displayName || "Unknown sender");
            } else {
              console.error("Failed to get sender:", fromResult.error);
              setSenderName("Unknown sender");
            }
          });
        } else {
          console.warn("Sender not available in this mode.");
          setSenderName("Unknown sender");
        }
      } else {
        // Read mode fallback
        const subjectVal = (item as Office.MessageRead).subject;
        const fromVal = (item as Office.MessageRead).from;

        console.log("Read mode detected.");
        console.log("Subject (read):", subjectVal);
        console.log("From (read):", fromVal);

        setSubject(typeof subjectVal === "string" ? subjectVal : "(No Subject)");
        setSenderName(fromVal?.displayName ?? "Unknown sender");
      }

      // Fetch PreVeil Build Version
      const fetchBuildVersion = async () => {
        try {
          const response = await fetch("/api/preveil/get/buildversion");
          if (!response.ok) throw new Error("Network response was not ok");

          const data = await response.json();
          console.log("✅ PreVeil Build Version:", data);
          setBuildVersion(data.version || JSON.stringify(data.version));
        } catch (error) {
          console.error("Error fetching build version:", error);
          setBuildVersion("Error fetching version");
        }
      };

      fetchBuildVersion();
    });
  }, []);

  const suggestedReply = `
    <p><br></p>
    <p>Hi World,</p>
    <p>Thanks for your message about "<strong>${subject}</strong>". I'll get back to you shortly.</p>
    <p>Best regards,</p>
  `;

  const handleInsert = () => {
    const item = Office.context?.mailbox?.item;
    if (item?.body?.setSelectedDataAsync) {
      item.body.setSelectedDataAsync(
        suggestedReply,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Reply successfully inserted.");
          } else {
            console.error("Insert failed:", result.error);
            alert("Could not insert the reply.");
          }
        }
      );
    } else {
      alert("Insert only works in compose mode.");
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h1>{title}</h1>
      <h2>Smart Reply Assistant</h2>

      <p><strong>From:</strong> {senderName}</p>
      <p><strong>Subject:</strong> {subject}</p>
      <p><strong>PreVeil Build Version:</strong> {buildVersion}</p>

      <TextField
        label="Suggested Reply"
        multiline
        readOnly
        value={suggestedReply}
        rows={6}
      />

      <PrimaryButton style={{ marginTop: 10 }} onClick={handleInsert}>
        Insert into Reply
      </PrimaryButton>
    </div>
  );
};

export default App;
