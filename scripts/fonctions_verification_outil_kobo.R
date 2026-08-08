################################################################################
# JMMI CAR - FONCTIONS DE VALIDATION DE L'OUTIL KOBO
################################################################################
# ENTREE      : kobo_survey, kobo_choices (tableaux deja charges par scripts/0_configuration.R)
# SORTIE      : check_constraints() renvoie un vecteur de caracteres pour chaque expression relevant
#               selected(question, 'reponse') dont la reponse n'est pas un choix reel de la liste de
#               cette question (vide si l'outil est coherent)
# DESCRIPTION : Analyse chaque expression de la colonne "relevant" contenant selected(...) et confirme
#               que l'option de reponse referencee existe bien dans la feuille des choix pour cette
#               question.
# OBJECTIF    : Detecter les fautes de frappe/renommages dans l'outil kobo lui-meme (ex. un relevant
#               referencant un nom de choix qui a ete modifie depuis) avant qu'ils ne cassent
#               silencieusement la logique en aval.
# CONTEXTE    : Porte depuis src/check_kobo.R de JMMI_ETH, avec qdapRegex::ex_default() remplace par un
#               petit equivalent base-R regmatches()/gregexpr() car qdapRegex n'est pas installe et
#               n'est necessaire pour rien d'autre que cet unique utilitaire d'extraction regex.
################################################################################

# Ces fonctions verifient la colonne "relevant" du questionnaire et signalent les problemes de
# conditions de type (selected(question_name, question_answer))

# Appeler cette fonction en lui passant le questionnaire et les choix ; le resultat est la liste des
# conditions incorrectes de type (selected(question_name, question_answer)), s'il y en a. Une erreur
# signifie que la reponse n'existe pas dans la feuille des choix

# NOTE : la version ETH de ce fichier utilisait qdapRegex::ex_default() pour extraire les
# correspondances regex. qdapRegex n'est pas installe ici et n'est pas necessaire : le base-R
# regmatches()/gregexpr() donne le meme resultat (une liste de vecteurs de caracteres, un element par
# chaine en entree), donc ex_default() ci-dessous est un petit remplacement local plutot qu'une
# nouvelle dependance.

#############################################################################################################

ex_default <- function(x, pattern) {
  regmatches(x, gregexpr(pattern, x, perl = TRUE))
}

# ------------------------------------------------------------------------------

check_answer_in_list <- function(questions, choices, constraint){

  if(!str_detect(constraint, ",")) return(T)

  question_regex <- "\\{([^()]+)\\}"
  answer_regex <- "\\'([^()]+)\\'"

  question <- gsub(question_regex, "\\1", str_extract_all(constraint, question_regex)[[1]])
  answer <- gsub(answer_regex, "\\1", str_extract_all(constraint, answer_regex)[[1]])

  question_type <- questions %>%
    filter(name==question) %>%
    filter(!grepl("^(begin|end)\\s+group$",type)) %>%
    pull(type)

  if (question_type=="calculate") return(T)
  else{
    listname <- gsub("^.*\\s", "", question_type)
    choices_list <- choices %>% filter(list_name==listname) %>% pull(name)
    return(answer %in% choices_list)
  }

}

# ------------------------------------------------------------------------------

check_constraints <- function(questions, choices){

  questions <- mutate_at(questions, c("name", "type"), ~str_trim(.))
  choices <- mutate_at(choices, c("list_name", "name"), ~str_trim(.))

  all_contraints <- questions %>% filter(grepl("selected",relevant)) %>% pull(relevant)
  all_contraints <- gsub('"',"'",all_contraints)
  rs_list <- map(all_contraints,~map_lgl(unlist(ex_default(.x, pattern = "selected\\s*\\([^\\)]*\\)")),
                                         ~check_answer_in_list(questions, choices, .)))

  map2(rs_list,seq_along(rs_list),
       ~if(length(which(!.x))!=0){
         return(unlist(ex_default(all_contraints[.y], pattern = "selected\\s*\\([^\\)]*\\)"))[which(!.x)])
       }) %>%
    unlist() %>% unique()
}
