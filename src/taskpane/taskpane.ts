/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    await waitForElement("folders-list");
    showFolders();
  }
});

//mostra a mensagem na tela para qual pagina foi
function showMessage(message) {
    const box = document.getElementById("message-box");
    if (!box) return;

    box.textContent = message;
    box.style.display = "block";

    // Fade-out depois de 3s
    setTimeout(() => {
        box.style.opacity = "1";
        box.style.transition = "opacity 0.8s";
        box.style.opacity = "0";

        setTimeout(() => {
            box.style.display = "none";
            box.style.opacity = "1";
        }, 800);
    }, 3000);
}
function waitForElement(id) {
  return new Promise(resolve => {
    const el = document.getElementById(id);
    if (el) return resolve(el);

    const obs = new MutationObserver(() => {
      const el = document.getElementById(id);
      if (el) {
        obs.disconnect();
        resolve(el);
      }
    });

    obs.observe(document.body, { childList: true, subtree: true });
  });
}

//pega o corpo do email
async function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
};

//pega os anexos do email
async function getAttachments(item) {
  if (!item.attachments || item.attachments.length === 0) {
    return [];
  }

  const attachmentsData = [];

  for (const att of item.attachments) {
    const content: any = await new Promise((resolve, reject) => {
  item.getAttachmentContentAsync(att.id, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      resolve(result.value as any);
    } else {
      reject(result.error);
    }
  });
});

    let contentType = att.contentType;
    if (!contentType) {
      const extension = att.name.split('.').pop().toLowerCase();
      contentType = getMimeType(extension) || "application/octet-stream";
    }

    attachmentsData.push({
      id: att.id,
      name: att.name,
      contentType: contentType,
      size: att.size,
      contentBytes: content.content, // Base64 OK
      isInline: att.isInline || false,
      attachmentType: att.attachmentType
    });
  }

  return attachmentsData;
}


// Helper para determinar MIME type
function getMimeType(extension) {
  const mimeTypes = {
    'pdf': 'application/pdf',
    'doc': 'application/msword',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'xls': 'application/vnd.ms-excel',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'ppt': 'application/vnd.ms-powerpoint',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'png': 'image/png',
    'gif': 'image/gif',
    'txt': 'text/plain',
    'zip': 'application/zip',
    'rar': 'application/x-rar-compressed'
  };
  
  return mimeTypes[extension] || null;
}

//Função que carrega as pastas do sharepoint, talvez depois usar o graph para ser mais rapido
async function carregarPastas() {
    try {
        const response = await fetch(url, { method: "POST" });

        if (!response.ok) {
            console.error("Erro ao buscar pastas:", response.statusText);
            return [];
        }

        const data = await response.json();
        console.log("🔵 RAW RESPONSE DO FLOW:", data);

        // O flow já retorna diretamente o array
        if (Array.isArray(data)) {
            console.log("🟢 Lista de pastas carregada:", data);
            return data;  // <-- aqui!
        }

        console.log("🔴 Formato inesperado:", data);
        return [];

    } catch (err) {
        console.error("Falha ao carregar pastas:", err);
        return [];
    }
}

//função que cria a pasta no sharepoint com o dominio do email
export async function run() {
  const item = Office.context.mailbox.item;

    // MOSTRA A SPLASH SCREEN
    const splash = document.getElementById("splash-screen")!;
    if (splash) splash.style.display = "flex";

    try {
    // Extrai informações básicas do e-mail
    const subject = item.subject;
    const from = item.from && item.from.emailAddress ? item.from.emailAddress : "";
    const body = await getEmailBody(item);
    const toRecipient = item.to && item.to.length > 0 ? item.to[0].emailAddress : "";
    const attachments = await getAttachments(item);

    // Monta o payload
    const data = {
      subject: subject,
      to: toRecipient,
      from: from,
      body: body,
      attachments: attachments
    };

    //link da automação que cria a pasta com o dominio de email de quem enviou o email
    const flowUrl = "https://defaulte8fc68b65d194bf4a2c1a5ed5dc4c2.f5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/147199a4a1cb4dbe98d5119cffa803bd/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2kLY1Qkb-zgjJuEIpGJBR94VBHYMV-qkPgel0ubfu_U"; // coloque aqui a URL do gatilho HTTP real
      const response = await fetch(flowUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(data)
      });

      if (splash) splash.style.display = "none";

      if (response.ok) {
        console.log("Fluxo acionado com sucesso!");
        alert("Fluxo iniciado com sucesso!");
      } else {
        console.error("Erro ao acionar fluxo:", response.statusText);
        alert("Erro ao acionar fluxo.");
      }
    } catch (error) {
      if (splash) splash.style.display = "none";
      console.error("Falha na requisição:", error);
      alert("Falha ao conectar ao Power Automate.");
    }
};

//mostra as pastas na tela
async function showFolders() {
  const container = document.getElementById("folders-list");
  if (!container) {
    console.error("Elemento #folders-list não encontrado no HTML.");
    return;
  }

  container.innerHTML = "<p>Carregando pastas...</p>";

  try {
    const folders = await carregarPastas();

    if (folders.length === 0) {
      container.innerHTML = "<p>Nenhuma pasta encontrada.</p>";
      return;
    }

    // Cria HTML para cada pasta, junto do ícone
    container.innerHTML = folders
        .map(
        (f) => `
          <div class="folder-item" data-id="${f.id}" 
              style="cursor:pointer; padding:8px; border:1px solid #ddd; margin-bottom:5px; border-radius:5px; display:flex; align-items:center; gap:8px;">
            
            <img src="../../assets/folder.png" 
                alt="folder" width="20" height="20" 
                style="pointer-events:none;" />

            <span>${f.nome}</span>
          </div>`
      )
      .join("");
    // Adiciona evento de clique para cada pasta
    document.querySelectorAll(".folder-item").forEach((el) => {
      el.addEventListener("click", async (e) => {
        const folderName = (e.target as HTMLElement).textContent?.trim() || ""; // Remove o ícone e espaços
        showConfirm(folderName);
        /*console.log("Clicou na pasta:", folderName);
        showMessage(`Pasta selecionada: ${folderName}`);

        // Aqui você pode chamar seu fluxo do Power Automate:
        await uploadToFolder(folderName);*/
      });
    });
  } catch (err) {
    console.error(err);
    container.innerHTML = "<p>Erro ao carregar as pastas.</p>";
  }
};

//link da automação que pega as pastas do flow
const url = "https://defaulte8fc68b65d194bf4a2c1a5ed5dc4c2.f5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/ed37f3d5436d4e928c3a7680cf95b076/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jech5xSQ8_ib0EV2vu2JblGA9KP1cOJJGNpBsWM_BRY";


//Função que envia o email para a pasta selecionada junto com os anexos
async function uploadToFolder(folderName) {
  const item = Office.context.mailbox.item;
  const attachments = await getAttachments(item);

  const data = {
    folderName: folderName,
    subject: item.subject,
    from: item.from.emailAddress,
    to: item.to && item.to.length > 0 ? item.to[0].emailAddress : "",
    attachments: attachments,
    body: await getEmailBody(item)
  };

  //link da automação que salva na pasta selecionada
  const flowUrl = "https://defaulte8fc68b65d194bf4a2c1a5ed5dc4c2.f5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0a1f02a5fb85469c9c9202ee125a044a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=g78uUy_e-iG9t-rLhoLAPB40pmt6NHEOI0y_z3sYXUA";

  const response = await fetch(flowUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data)
  });
  

  /*if (response.ok) {
    showMessage(`Pasta criada com sucesso!`);
  } else {
    showMessage(`Erro ao criar pasta.`);
  }*/
}

function showConfirm(folderName: string) {
  const modal = document.getElementById("confirm-modal")!;
  const text = document.getElementById("confirm-text")!;
  const splash = document.getElementById("splash-screen")!;

  // Mensagem personalizada
  text.textContent = `Deseja salvar o e-mail na pasta "${folderName}"?`;

  modal.style.display = "flex";

  // Botões
  const btnYes = document.getElementById("btn-confirm-yes")!;
  const btnNo = document.getElementById("btn-confirm-no")!;

  // Remove eventos antigos para evitar duplicações
  btnYes.replaceWith(btnYes.cloneNode(true));
  btnNo.replaceWith(btnNo.cloneNode(true));

  const newYes = document.getElementById("btn-confirm-yes")!;
  const newNo = document.getElementById("btn-confirm-no")!;

  newYes.addEventListener("click", async () => {
    modal.style.display = "none";

    // MOSTRA A SPLASH SCREEN
    splash.style.display = "flex";

    try {
      await uploadToFolder(folderName); // <-- chama seu flow

      // ESCONDE SPLASH
      splash.style.display = "none";

      showMessage("E-mail salvo na pasta com sucesso!");
    } catch (err) {
      splash.style.display = "none";
      showMessage("Ocorreu um erro ao enviar.");
      console.error(err);
    }
  });

  newNo.addEventListener("click", () => {
    modal.style.display = "none";
    showMessage("Cancelado.");
  });
}