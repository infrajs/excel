<h1 class="text-danger">Excel-документ {data.src}</h1>
<a class="float-right" href="/~price.xlsx">Скачать</a>

{data.childs::list}

{list:}
	<h2>{title}</h2>
	<ul>
		{descr::descr}
	</ul>
	<table class="table table-striped">
		<tr>
			{head::head}
		</tr>
		{data::row}
	</table>
{row:}
	<tr>
		{~obj(:key,~key,:column,...head,:poss,...data).column::cell}
	</tr>
	{cell:}
		<td>{...poss[...key][.]:op}</td>
	{op:}{.}
{head:}
	<th>{.}</th>
{descr:}
	<li><b>{~key}</b>: {.}</li>
