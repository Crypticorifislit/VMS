document.addEventListener('DOMContentLoaded', function() {
    // Initialize search with debouncing
    const searchInputs = document.querySelectorAll('.filter-group input[type="text"]');
    searchInputs.forEach(input => {
        input.addEventListener('input', debounce(handleSearch, 600));
    });

    // Initialize per-page select
    const perPageSelect = document.querySelector('.per-page-select');
    if (perPageSelect) {
        perPageSelect.addEventListener('change', handlePerPageChange);
    }
});

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function handleSearch(event) {
    const form = event.target.closest('form');
    form.submit();
}

function handlePerPageChange(event) {
    const url = new URL(window.location);
    url.searchParams.set('per_page', event.target.value);
    url.searchParams.set('page', 1); // Reset to first page when changing items per page
    window.location = url;
}

let organizationToDelete = null;

function deleteOrganization(organizationId) {
    organizationToDelete = organizationId;
    const modal = document.getElementById('deleteModal');
    modal.style.display = 'block';
}

document.getElementById('cancelDelete').addEventListener('click', function() {
    document.getElementById('deleteModal').style.display = 'none';
    organizationToDelete = null;
});

document.getElementById('confirmDelete').addEventListener('click', async function() {
    if (organizationToDelete) {
        try {
            const response = await fetch(`/organizations/delete/${organizationToDelete}`, {
                method: 'DELETE',
            });
            
            if (response.ok) {
                window.location.reload();
            } else {
                const data = await response.json();
                alert(data.error || 'Error deleting organization');
            }
        } catch (error) {
            alert('Error deleting organization');
        }
    }
    document.getElementById('deleteModal').style.display = 'none';
});

// Close modal when clicking outside
window.addEventListener('click', function(event) {
    const modal = document.getElementById('deleteModal');
    if (event.target === modal) {
        modal.style.display = 'none';
        organizationToDelete = null;
    }
});
