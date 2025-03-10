        document.addEventListener("DOMContentLoaded", 
        function()
        {
            var chooser = document.getElementById("previewSelect");
            var initial = document.getElementById("target-preview");

            var initialHref = initial.href.substring(initial.href.lastIndexOf("/") + 1);
            chooser.value = initialHref;

            chooser.addEventListener("change", 
            function()
            {
                var selected = this.value;
                var current = document.getElementById("target-preview");

                current.setAttribute("href", selected);
            }
        )})