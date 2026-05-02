from django.shortcuts import render

from .services import load_documents


def document_list(request):
    all_documents = load_documents()
    documents = all_documents

    keyword = request.GET.get("q", "").strip()

    if keyword:
        filtered_documents = []

        for doc in documents:
            target_text = " ".join([
                str(doc.get("id", "")),
                str(doc.get("category", "")),
                str(doc.get("document_name", "")),
                str(doc.get("required_level", "")),
                str(doc.get("owner_dept", "")),
                str(doc.get("status", "")),
            ])

            if keyword.lower() in target_text.lower():
                filtered_documents.append(doc)

        documents = filtered_documents

    return render(request, "documents/document_list.html", {
        "documents": documents,
        "keyword": keyword,
        "total_count": len(all_documents),
        "display_count": len(documents),
    })