# Applied to the raw Actor output before committing:
#   jq '[.[] | select(.id == "dQw4w9WgXcQ") | .comments |= .[0:15] | .commentsCount = (.comments | length)]' out/raw.json \
#     | jq -f samples/redact.jq > samples/rick-astley-top-comments.json
#
# The first command keeps one video and its first 15 top comments (the run returned 1,000) with
# their replies, and sets commentsCount to match. This file then pseudonymizes every commenter:
# - Each comment or reply author becomes @commenter_01, @commenter_02 ... (in order of first
#   appearance), verified or not. Their channel id and avatar URL are set to null.
# - Comment ids become comment_01, comment_02 ...; a reply's id becomes <parent alias>.reply_N.
# - Comment text is cut to its first sentence (plus "…" when something was cut). Commenters'
#   handles anywhere in the text become their alias (case-insensitive, with or without the @),
#   except verified accounts' handles (a brand name such as "YouTube" stays a word), other
#   @mentions become @user, emails are removed.
# The video (id, title, URL, channel, counts, description) and every count, date and flag of each
# comment are left untouched.

def rx: gsub("(?<c>[.*+?^${}()|\\[\\]\\\\/-])"; "\\\(.c)");

def firstSentence:
  (gsub("\\s+"; " ") | sub("^ "; "") | sub(" $"; "")) as $t
  | ([$t | capture("^(?<s>.+?[.!?。！？])( |$)")] | first | .s?) as $s
  | if $s != null and ($s | length) <= 300 then ($s + (if ($s | length) < ($t | length) then " …" else "" end))
    elif ($t | split(" ") | length) > 25 then ($t | split(" ")[:25] | join(" ")) + " …"
    else $t end;

def pad($n): ($n | tostring | if length < 2 then "0" + . else . end);
def bare: sub("^@"; "");

# Every author, comments and replies, in document order, without duplicates.
([.[] | .comments[]? | (., (.replies // [])[]) | .author | select(type == "string")]
  | reduce .[] as $a ([]; if index([$a]) then . else . + [$a] end)) as $authors
| ([.[] | .comments[]? | (., (.replies // [])[]) | select(.authorIsVerified == true) | .author] | unique) as $verified
| ([.[] | .comments[]? | .commentId]) as $ids
| def alias($a): "commenter_" + pad(($authors | index([$a])) + 1);
  def cid($id): "comment_" + pad(($ids | index([$id])) + 1);
  # \u0001 marks an aliased @mention so the generic @mention rule below leaves it alone.
  def scrub:
    if type == "string" then
      (reduce ($authors - $verified)[] as $x (.; ($x | bare) as $h
          | if ($h | length) > 3 then
              gsub("(?<![A-Za-z0-9_.-])@" + ($h | rx) + "(?![\\w-])"; "\u0001" + alias($x); "i")
              | gsub("(?<![A-Za-z0-9_.@\u0001-])" + ($h | rx) + "(?![\\w-])"; alias($x); "i")
            else . end))
      | gsub("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}"; "(email removed)")
      | gsub("@[\\w.\\-]+"; "@user")
      | gsub("\u0001"; "@")
      | firstSentence
    else . end;
  def person: .author |= (if type == "string" then "@" + alias(.) else . end)
    | .authorChannelId = null | .authorAvatarUrl = null | .text |= scrub;
  [ .[] | if (.comments | type) == "array" then
      .comments |= map(
        cid(.commentId) as $c
        | person | .commentId = $c
        | if (.replies | type) == "array"
          then .replies |= (to_entries | map(.key as $i | .value | person | .commentId = "\($c).reply_\($i + 1)"))
          else . end)
    else . end ]
